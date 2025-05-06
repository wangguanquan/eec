package org.ttzero.excel.entity;

import org.ttzero.excel.annotation.NestedObject;
import org.ttzero.excel.reader.Cell;
import org.ttzero.excel.util.StringUtil;

import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import static org.ttzero.excel.util.ReflectUtil.listDeclaredFieldsUntilJavaPackage;

/**
 * 支持嵌套对象（@NestedObject）的 ListSheet。
 * 递归展开所有嵌套对象字段，自动合并为表头列。
 * 兼容 ListSheet 的全部特性。
 * @author Chai at 2025/4/15 14:38
 */
public class NestedListSheet<T> extends ListSheet<T> {

    private String columnNameFormat;

    public NestedListSheet() {
        super();
    }

    public NestedListSheet(String name) {
        super(name);
    }

    public NestedListSheet(final Column... columns) {
        super(columns);
    }

    public NestedListSheet(String name, final Column... columns) {
        super(name, columns);
    }

    public NestedListSheet(List<T> data) {
        super(data);
    }

    public NestedListSheet(String name, List<T> data) {
        super(name, data);
    }

    public NestedListSheet(List<T> data, final Column... columns) {
        super(data, columns);
    }

    public NestedListSheet(String name, List<T> data, final Column... columns) {
        super(name, data, columns);
    }

    public NestedListSheet(List<T> data, WaterMark waterMark, final Column... columns) {
        super(data, waterMark, columns);
    }

    public NestedListSheet(String name, List<T> data, WaterMark waterMark, final Column... columns) {
        super(name, data, waterMark, columns);
    }

    /**
     * 设置列名格式
     * @param columnNameFormat 列名格式
     * @return this
     */
    protected NestedListSheet setColumnNameFormat(String columnNameFormat) {
        if (StringUtil.isBlank(columnNameFormat)) {
            return this;
        }
        if (StringUtil.isBlank(this.columnNameFormat)) {
            this.columnNameFormat = columnNameFormat;
        } else {
            this.columnNameFormat = columnNameFormat + this.columnNameFormat.replace("%s", "");
        }
        return this;
    }

    /**
     * 初始化表头，递归展开所有 @NestedObject 字段。
     * 先调用父类 ListSheet 的 init，再补充嵌套对象列。
     * @return 列总数
     */
    @Override
    protected int init() {
        if (hasHeaderColumns()) {
            return columns.length;
        }
        int count = super.init();
        Class<?> clazz = getTClass();
        if (clazz == null) return columns != null ? columns.length : 0;
        // 收集所有带 @NestedObject 的字段，递归展开
        Field[] declaredFields = listDeclaredFieldsUntilJavaPackage(clazz, f -> f.isAnnotationPresent(NestedObject.class));
        if (declaredFields.length == 0) {
            return count;
        }
        List<Column> nestedColumns = new ArrayList<>();
        for (Field field : declaredFields) {
            if (!field.isAccessible()) field.setAccessible(true);
            NestedObject nestedObject = field.getAnnotation(NestedObject.class);

            // 递归生成嵌套 sheet 并展开其列
            NestedListSheet<?> nestedSheet = new NestedListSheet<>();
            nestedSheet.setColumnNameFormat(nestedObject.columnNameFormat());
            if (nestedObject.columnNameFormatExtend()) {
                nestedSheet.setColumnNameFormat(columnNameFormat);
            }
            nestedSheet.setClass(field.getType());
            nestedSheet.init();

            Column[] nestedCols = nestedSheet.columns;
            int startColIndex = nestedObject.startColIndex();
            if (nestedCols != null) {
                for (Column col : nestedCols) {
                    if (nestedSheet.columnNameFormat != null && ((EntryColumn) col).rootFields == null) {
                        col.name = String.format(nestedSheet.columnNameFormat, col.name);
                    }
                    if (startColIndex > 0) {
                        col.colIndex += startColIndex;
                    }
                    // 设置 rootField
                    if (col instanceof EntryColumn) {
                        ((EntryColumn) col).addFirstRootField(field);
                    }
                    nestedColumns.add(col);
                }
            }
        }
        // 合并父类 columns 和嵌套 columns
        List<Column> merged = new ArrayList<>(Arrays.asList(columns));
        merged.addAll(nestedColumns);
        columns = merged.toArray(new Column[0]);
        return columns.length;
    }

    /**
     * 重置 RowBlock 行块数据，支持嵌套对象的递归取值。
     * 普通字段逻辑完全复用父类。
     */
    @Override
    protected void resetBlockData() {
        if (!eof && left() < rowBlock.capacity()) {
            append();
        }
        int end = getEndIndex();
        int len = columns.length;
        boolean hasGlobalStyleProcessor = (extPropMark & 2) == 2;
        try {
            for (; start < end; rows++, start++) {
                Row row = rowBlock.next();
                row.index = rows;
                Cell[] cells = row.realloc(len);
                T o = data.get(start);
                boolean isNull = o == null;
                for (int i = 0; i < len; i++) {
                    Cell cell = cells[i];
                    cell.clear();
                    EntryColumn column = (EntryColumn) columns[i];
                    // 递归获取嵌套对象的最终值
                    Object e = getNestedOrDirectValue(o, column, isNull);
                    cellValueAndStyle.reset(row, cell, e, column);
                    if (hasGlobalStyleProcessor) {
                        cellValueAndStyle.setStyleDesign(o, cell, column, getStyleProcessor());
                    }
                }
                row.height = getRowHeight();
            }
        } catch (IllegalAccessException | java.lang.reflect.InvocationTargetException e) {
            throw new ExcelWriteException(e);
        }
    }

    /**
     * 递归获取嵌套字段的最终对象。
     * 判断 column 是否为嵌套对象字段，通过 field 的类型是否为复杂对象（非基本类型/非 String/非包装类）来判断。
     * @param root   根对象
     * @param column 当前 EntryColumn
     * @param isNull 根对象是否为 null
     * @return 字段/方法的最终值
     */
    protected Object getNestedOrDirectValue(Object root, EntryColumn column, boolean isNull) throws IllegalAccessException, InvocationTargetException {
        if (isNull || column == null) return null;
        Field[] rootFields = column.rootFields;
        if (rootFields != null && rootFields.length > 0) {
            Field lastRootField = rootFields[rootFields.length - 1];
            if (!lastRootField.getType().equals(root.getClass())) {
                Object parentObj = root;
                for (Field rootField : rootFields) {
                    parentObj = rootField.get(parentObj);
                    if (parentObj == null) break;
                }
                return getNestedOrDirectValue(parentObj, column, parentObj == null);
            }
        }
        // 普通 getter/field
        if (column.method != null) {
            return column.method.invoke(root);
        }
        if (column.field != null) {
            return column.field.get(root);
        }
        return root;
    }

}

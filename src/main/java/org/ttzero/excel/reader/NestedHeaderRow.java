package org.ttzero.excel.reader;

import org.ttzero.excel.annotation.NestedObject;
import org.ttzero.excel.entity.ListSheet;
import org.ttzero.excel.util.StringUtil;

import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.util.Comparator;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;
import java.util.stream.Stream;

import static org.ttzero.excel.util.ReflectUtil.listDeclaredFieldsUntilJavaPackage;

/**
 * 嵌套对象行数据解析
 * @author Chai at 2025/4/7 9:35
 */
public class NestedHeaderRow extends HeaderRow {
    /**
     * 表头包含匹配选项
     */
    public static final int CONTAINS_HEADER = 1 << 3;
    /**
     * 空嵌套对象设为null选项
     */
    public static final int NULLIFY_WHEN_ALL_NESTED_FIELDS_EMPTY = 1 << 4;

    private Map<Field, NestedHeaderRow> nestedColumnMap;
    private NestedObject nestedObject;
    private String columnNameFormat;
    private Field rootField;

    public NestedHeaderRow() {
    }

    public NestedHeaderRow(NestedObject nestedObject) {
        this(nestedObject, null);
    }

    public NestedHeaderRow(NestedObject nestedObject, Field rootField) {
        this.nestedObject = nestedObject;
        this.rootField = rootField;
        setColumnNameFormat(nestedObject.columnNameFormat());
    }

    /**
     * 获取列名对应的索引位置
     */
    @Override
    public int getIndex(String columnName) {
        if (mapping == null) {
            return -1;
        }
        if ((option & 2) == 2) columnName = columnName.toLowerCase();
        if (StringUtil.isNotBlank(columnNameFormat)) columnName = String.format(columnNameFormat, columnName);

        Integer index = null;
        if ((option & 8) == 8) {
            List<Map.Entry<String, Integer>> entries = mapping.entrySet().stream()
                    .sorted(Map.Entry.comparingByValue())
                    .collect(Collectors.toList());
            for (Map.Entry<String, Integer> entry : entries) {
                if (entry.getKey().contains(columnName)) {
                    index = entry.getValue();
                    break;
                }
            }
        } else {
            index = mapping.get(columnName);
        }
        return index != null ? index : -1;
    }

    /**
     * 设置关联的Java类并解析其字段和嵌套对象结构
     */
    @Override
    protected HeaderRow setClass(Class<?> clazz) {
        super.setClass(clazz);
        Field[] declaredFields = listDeclaredFieldsUntilJavaPackage(clazz, f -> f.isAnnotationPresent(NestedObject.class));

        if (declaredFields.length > 0) {
            if (nestedColumnMap == null) {
                nestedColumnMap = new HashMap<>(declaredFields.length);
            }
            for (Field f : declaredFields) {
                if (!f.isAccessible()) f.setAccessible(true);
                NestedObject object = f.getAnnotation(NestedObject.class);
                NestedHeaderRow headerRow = new NestedHeaderRow(object, f);
                headerRow.with(this).setOptions(option << 16 >>> 16);
                // 拼接父类列名格式化规则
                if (object.columnNameFormatExtend()) headerRow.setColumnNameFormat(columnNameFormat);
                headerRow.setClass(f.getType());
                for (ListSheet.EntryColumn column : headerRow.columns) {
                    if (rootField != null) column.addFirstRootField(rootField);
                    column.addFirstRootField(f);
                    column.setName(String.format(headerRow.columnNameFormat, column.getName()));
                }
                nestedColumnMap.put(f, headerRow);
            }
        }

        if (nestedObject != null && nestedObject.startColIndex() >= 0 && columns != null) {
            int startColIndex = nestedObject.startColIndex();
            for (ListSheet.EntryColumn column : this.columns) {
                if (column.colIndex >= 0) {
                    column.colIndex += startColIndex;
                }
            }
        }
        return this;
    }

    /**
     * 获取所有列信息（包括嵌套对象中的列）
     * @return 列信息
     */
    protected ListSheet.EntryColumn[] getAllColumns() {
        ListSheet.EntryColumn[] columns = this.columns;
        if (nestedColumnMap == null || nestedColumnMap.isEmpty()) {
            return columns;
        }
        List<ListSheet.EntryColumn> nestedColumnList = nestedColumnMap.values().stream()
                .map(NestedHeaderRow::getAllColumns)
                .flatMap(Stream::of)
                .collect(Collectors.toList());
        if (columns == null) {
            return nestedColumnList.toArray(new ListSheet.EntryColumn[0]);
        }
        nestedColumnList.addAll(Stream.of(columns).collect(Collectors.toList()));
        nestedColumnList.sort(Comparator.comparingInt(column -> column.colIndex));
        return nestedColumnList.toArray(new ListSheet.EntryColumn[0]);
    }

    /**
     * 设置列名格式
     * @param columnNameFormat 列名格式
     * @return this
     */
    protected NestedHeaderRow setColumnNameFormat(String columnNameFormat) {
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
     * 判断此嵌套对象中的所有字段是否为空
     * @param row 行数据
     * @return 是否为空
     */
    protected boolean isAllColumnValueBlank(Row row) {
        if ((option & 16) == 0) {
            return false;
        }
        if (columns == null) {
            return true;
        }
        for (ListSheet.EntryColumn ec : columns) {
            if (ec.colIndex > 0 && !row.isBlank(ec.colIndex)) {
                return false;
            }
        }
        return true;
    }

    @Override
    void put(Row row, Object t) throws IllegalAccessException, InvocationTargetException {
        super.put(row, t);
        if (nestedColumnMap == null || nestedColumnMap.isEmpty()) {
            return;
        }
        Field field = null;
        for (Map.Entry<Field, NestedHeaderRow> entry : nestedColumnMap.entrySet()) {
            field = entry.getKey();
            NestedHeaderRow headerRow = entry.getValue();
            if (headerRow.isAllColumnValueBlank(row)) {
                field.set(t, null);
            } else {
                try {
                    Object nestedObject = field.getType().newInstance();
                    field.set(t, nestedObject);
                    headerRow.put(row, nestedObject);
                } catch (InstantiationException e) {
                    throw new RuntimeException("Unable to create nested object instance: {" + field.getType() + "}");
                }
            }
        }
    }
}

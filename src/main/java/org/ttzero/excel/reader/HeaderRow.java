/*
 * Copyright (c) 2017-2019, guanquan.wang@yandex.com All Rights Reserved.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

package org.ttzero.excel.reader;

import org.ttzero.excel.annotation.ExcelColumn;
import org.ttzero.excel.annotation.ExcelColumns;
import org.ttzero.excel.annotation.IgnoreImport;
import org.ttzero.excel.annotation.RowNum;
import org.ttzero.excel.entity.Column;
import org.ttzero.excel.entity.ListSheet;
import org.ttzero.excel.manager.Const;
import org.ttzero.excel.util.StringUtil;

import java.beans.IntrospectionException;
import java.beans.Introspector;
import java.beans.PropertyDescriptor;
import java.lang.annotation.Annotation;
import java.lang.reflect.AccessibleObject;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.math.BigDecimal;
import java.time.DateTimeException;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.Comparator;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.StringJoiner;
import java.util.stream.Collectors;

import static org.ttzero.excel.entity.IWorksheetWriter.isBool;
import static org.ttzero.excel.entity.IWorksheetWriter.isChar;
import static org.ttzero.excel.entity.IWorksheetWriter.isDate;
import static org.ttzero.excel.entity.IWorksheetWriter.isInt;
import static org.ttzero.excel.entity.IWorksheetWriter.isLocalDate;
import static org.ttzero.excel.entity.IWorksheetWriter.isLocalDateTime;
import static org.ttzero.excel.entity.IWorksheetWriter.isLocalTime;
import static org.ttzero.excel.entity.Sheet.int2Col;
import static org.ttzero.excel.util.ReflectUtil.listDeclaredFields;
import static org.ttzero.excel.util.ReflectUtil.listDeclaredMethods;
import static org.ttzero.excel.util.StringUtil.EMPTY;
import static org.ttzero.excel.util.StringUtil.lowFirstKey;
import static org.ttzero.excel.util.StringUtil.toCamelCase;

/**
 * @author guanquan.wang at 2019-04-17 11:55
 */
public class HeaderRow extends Row {
    /**
     * Force Import (Match with field name if without {@code ExcelColumn} annotation)
     */
    public static final int FORCE_IMPORT = 1;
    /**
     * Ignore Case (Ignore case matching column names)
     */
    public static final int  IGNORE_CASE = 1 << 1;
    /**
     * Camel Case (CAMEL_CASE to camelCase)
     */
    public static final int  CAMEL_CASE = 1 << 2;

    protected String[] names;
    protected Class<?> clazz;
    protected Object t;
    /* The column name and column position mapping */
    protected Map<String, Integer> mapping;
    /* Storage header column */
    protected ListSheet.EntryColumn[] columns;

    // `detailMessage` field declare in Throwable
    protected static final Field detailMessageField;
    // Specify total rows of header
    protected int headRows;
    /**
     * Simple properties
     *
     * <blockquote><pre>
     *  Bit  | Contents
     * ------+---------
     * 31. 1 | Force Import
     * 30. 1 | Ignore Case (Ignore case matching column names)
     * 29. 1 | Camel Case
     * </pre></blockquote>
     */
    protected int option;

    static {
        Field field = null;
        try {
            field = Throwable.class.getDeclaredField("detailMessage");
            field.setAccessible(true);
        } catch (Exception e) {
            // Ignore
        }
        detailMessageField = field;
    }

    public HeaderRow with(Row ... rows) {
        return with(null, rows.length, rows);
    }

    public HeaderRow with(int headRows, Row ... rows) {
        return with(null, headRows, rows);
    }

    public HeaderRow with(List<Dimension> mergeCells, Row ... rows) {
        return with(mergeCells, rows.length, rows);
    }

    public HeaderRow with(List<Dimension> mergeCells, int headRows, Row ... rows) {
        this.headRows = headRows;
        Row row = rows[rows.length - 1];
        if (row == null) return new HeaderRow();
        this.names = new String[row.lc];
        this.mapping = new HashMap<>();
        // Extends from row
        this.fc = row.fc;
        this.lc = row.lc;
        this.index = row.index;
        this.cells = new Cell[this.names.length];
        for (int i = 0; i < row.fc; i++) {
            this.cells[i] = new Cell();
        }

        if (headRows == 1) {
            for (int i = row.fc; i < row.lc; i++) {
                this.names[i] = row.getString(i);
                this.mapping.put(makeKey(this.names[i]), i);

                Cell cell = new Cell();
                cell.setSv(this.names[i]);
                this.cells[i] = cell;
            }
        } else {
            // Copy on merge cells
            mergeCellsIfNull(mergeCells, rows);

            StringBuilder buf = new StringBuilder();
            for (int i = row.fc; i < row.lc; i++) {
                buf.delete(0, buf.length());
                for (Row r : rows) {
                    String tmp = r.getString(i);
                    if (StringUtil.isNotEmpty(tmp)) {
                        buf.append(tmp).append(':');
                    }
                }
                if (buf.length() > 1) buf.deleteCharAt(buf.length() - 1);
                this.names[i] = buf.toString();
                this.mapping.put(makeKey(this.names[i]), i);

                Cell cell = new Cell();
                cell.setSv(this.names[i]);
                this.cells[i] = cell;
            }
        }
        return this;
    }

    public final boolean is(Class<?> clazz) {
        return this.clazz != null && this.clazz == clazz;
    }

    /**
     * mapping
     *
     * @param clazz the type of binding
     * @return the header row
     */
    protected HeaderRow setClass(Class<?> clazz) {
        this.clazz = clazz;
        // Parse Field
        Field[] declaredFields = listDeclaredFields(clazz, c -> !ignoreColumn(c));

        // Parse Method
        Map<String, Method> tmp = new HashMap<>();
        try {
            PropertyDescriptor[] propertyDescriptors = Introspector.getBeanInfo(clazz, Object.class)
                    .getPropertyDescriptors();
            for (PropertyDescriptor pd : propertyDescriptors) {
                Method method = pd.getWriteMethod();
                if (method != null) tmp.put(pd.getName(), method);
            }
        } catch (IntrospectionException e) {
            LOGGER.warn("Get class {} methods failed.", clazz);
        }

        List<ListSheet.EntryColumn> list = new ArrayList<>();
//        Map<String, ListSheet.EntryColumn> columnMap = new LinkedHashMap<>();
        ListSheet.EntryColumn column;
        for (int i = 0; i < declaredFields.length; i++) {
            Field f = declaredFields[i];
            f.setAccessible(true);
            String gs = f.getName();

            // The setter methods take precedence over property reflection
            Method method = tmp.get(gs);
            if (method != null) {
                column = createColumn(method);
                if (column != null) {
                    ListSheet.EntryColumn tail = column.tail != null ? (ListSheet.EntryColumn) column.tail : column;
                    tail.method = method;
                    if (StringUtil.isEmpty(tail.name)) tail.name = method.getName();
                    if (tail.clazz == null) tail.clazz = method.getParameterTypes()[0];
                    if (tail.colIndex < 0) tail.colIndex = check(tail.name, gs);
//                    if ((other = columnMap.get(tail.getName())) == null || other.getMethod() == null) columnMap.put(tail.name, column);
                    list.add(column);
                    continue;
                }
            }

            column = createColumn(f);
            if (column != null) {
                ListSheet.EntryColumn tail = column.tail != null ? (ListSheet.EntryColumn) column.tail : column;
                if (StringUtil.isEmpty(tail.name)) tail.name = gs;
                if (method != null) {
                    tail.method = method;
                    if (tail.clazz == null) tail.clazz = method.getParameterTypes()[0];
                } else {
                    tail.field = f;
                    if (tail.clazz == null) tail.clazz = declaredFields[i].getType();
                }
                if (tail.colIndex < 0) tail.colIndex = check(tail.name, gs);
//                if ((other = columnMap.get(tail.getName())) == null || other.getMethod() == null) columnMap.put(tail.name, column);
                list.add(column);
            }
        }

        // Others
        Map<String, Method> otherColumns = attachOtherColumn(clazz);

        if (!otherColumns.isEmpty()) {
            for (Map.Entry<String, Method> entry : otherColumns.entrySet()) {
                column = createColumn(entry.getValue());
                if (column == null) column = new ListSheet.EntryColumn(entry.getKey());
                ListSheet.EntryColumn tail = column.tail != null ? (ListSheet.EntryColumn) column.tail : column;
                if (StringUtil.isEmpty(tail.name)) tail.name = entry.getKey();
                tail.method = entry.getValue();
                if (tail.clazz == null) tail.clazz = entry.getValue().getParameterTypes()[0];
                if (tail.colIndex < 0) tail.colIndex = getIndex(tail.name);
//                if ((other = columnMap.get(tail.getName())) == null || other.getMethod() == null) columnMap.put(tail.name, column);
                list.add(column);
            }
        }

        // Merge cells
        org.ttzero.excel.entity.Sheet listSheet = new ListSheet<Object>() {
            @Override
            public Column[] getAndSortHeaderColumns() {
                columns = new Column[list.size()];
                list.toArray(columns);
                headerReady = true;

                // Sort column index
                sortColumns(columns);

                // Turn to one-base
                calculateRealColIndex();

                // Reverse
                reverseHeadColumn();

                // Add merge cell properties
                mergeHeaderCellsIfEquals();

                return columns;
            }
        };
        Column[] columns = listSheet.getAndSortHeaderColumns();
        @SuppressWarnings("unchecked")
        List<Dimension> mergeCells = (List<Dimension>) listSheet.getExtPropValue(Const.ExtendPropertyKey.MERGE_CELLS);
        // Ignore all vertical merged cells
        mergeCells = mergeCells != null && !mergeCells.isEmpty() ? mergeCells.stream().filter(c -> c.firstColumn < c.lastColumn).collect(Collectors.toList()) : null;
        if (mergeCells != null && !mergeCells.isEmpty()) {
            Grid mergedGrid = GridFactory.create(mergeCells);
            for (Column c : columns) {
                Column[] sub = c.toArray();
                for (int j = sub.length - 1, t; j >= 0; j--) {
                    if (mergedGrid.test(t = sub.length - j, c.realColIndex) && isTopRow(mergeCells, t, c.realColIndex)) {
                        Cell cell = new Cell((short) c.realColIndex);
                        if (StringUtil.isNotEmpty(sub[j].getName())) cell.setSv(sub[j].getName());
                        else cell.t = Cell.EMPTY_TAG;
                        mergedGrid.merge(t, cell);
                        sub[j].name = cell.sv;
                    }
                }
            }
        }

        for (int i = 0, len = columns.length; i < len; i++) {
            Column c = columns[i];
            if (c.tail != null) {
                StringJoiner joiner = new StringJoiner(":");
                Column[] sub = c.toArray();
                for (int j = sub.length - 1; j >= 0; j--) {
                    if (StringUtil.isNotEmpty(sub[j].getName())) joiner.add(sub[j].getName());
                }
                c.name = joiner.toString();
            } else if (!(c instanceof ListSheet.EntryColumn)) {
                columns[i] = c = new ListSheet.EntryColumn(c);
            }

            if (c.colIndex < 0) {
                c.colIndex = getIndex(c.name);
                c.realColIndex = c.colIndex + 1;
            }
        }

        this.columns = Arrays.stream(columns)
                .filter(c -> c.colIndex >= 0 || c.clazz == RowNum.class)
                .sorted(Comparator.comparingInt(a -> a.colIndex))
                .map(e -> (e instanceof ListSheet.EntryColumn) ? (ListSheet.EntryColumn) e : new ListSheet.EntryColumn(e))
                .toArray(ListSheet.EntryColumn[]::new);

        return this;
    }

    static boolean isTopRow(List<Dimension> mergeCells, int row, int col) {
        for (Dimension dim : mergeCells) {
            if (dim.checkRange(row, col) && row == dim.firstRow) return true;
        }
        return false;
    }

    protected int check(String first, String second) {
        int n = getIndex(first);
        if (n == -1) n = getIndex(second);
        if (n == -1) {
            LOGGER.warn("{} field [{}] can't find in header {}", clazz, first, Arrays.toString(names));
        }
        return n;
    }

    /**
     * mapping and instance
     *
     * @param clazz the type of binding
     * @return the header row
     * @throws IllegalAccessException -
     * @throws InstantiationException -
     */
    protected HeaderRow setClassOnce(Class<?> clazz) throws IllegalAccessException, InstantiationException {
        setClass(clazz);
        this.t = clazz.newInstance();
        return this;
    }

    protected ListSheet.EntryColumn[] getColumns() {
        return columns;
    }

    @SuppressWarnings("unchecked")
    public <T> T getT() {
        return (T) t;
    }

    public Class<?> getClazz() {
        return clazz;
    }

    @Override
    public CellType getCellType(int columnIndex) {
        return CellType.STRING;
    }

    /**
     * Get the column name by column index
     *
     * @param columnIndex the cell index
     * @return name of column
     */
    public String get(int columnIndex) {
        rangeCheck(columnIndex);
        return names[columnIndex];
    }

    /**
     * Get the name of columns
     *
     * @return name of columns
     */
    public String[] getNames() {
        return this.names;
    }

    /**
     * Returns the position in cell range
     *
     * @param columnName the column name
     * @return the position if found otherwise -1
     */
    public int getIndex(String columnName) {
        if (mapping != null) {
            if ((option & 2) == 2) columnName = columnName.toLowerCase();
            Integer index = mapping.get(columnName);
            return index != null ? index : -1;
        }
        return -1;
    }

    @Override
    public String toString() {
        StringJoiner joiner = new StringJoiner(" | ");
        StringBuilder buf = new StringBuilder();
        int i = 0;
        for (; i < names.length && names[i] == null; i++) ;
        char[] chars = new char[10];
        Arrays.fill(chars, 0, chars.length, '-');
        for (int j = i; i < names.length; i++) {
            joiner.add(names[i]);
            int n = simpleTestLength(names[i]) + (j == i || i == names.length - 1 ? 1 : 2);
            if (n > chars.length) chars = new char[n];
            Arrays.fill(chars, 0, n, '-');

            if (columns != null && i < columns.length && columns[i].clazz != RowNum.class) {
                Class<?> c = columns[i].clazz;

                // Align Center
                if (isDate(c) || isLocalDate(c) || isLocalDateTime(c) || isLocalTime(c) || isChar(c) || isBool(c)) {
                    chars[0] = chars[n - 1] = ':';
                }
                // Align Right
                else if (isInt(c)) {
                    chars[n - 1] = ':';
                }
                // Align Left
//            else;
            }
            buf.append(chars, 0, n).append('|');

        }

        buf.insert(0, joiner.toString() + Const.lineSeparator);

        return buf.toString();
    }

    protected int simpleTestLength(String name) {
        if (name == null) return 4;
        char[] chars = name.toCharArray();
        double d = 0.0;
        for (char c : chars) {
            if (c < 0x80) d++;
            else d += 1.75;
        }
        return (int) d;
    }

    void put(Row row, Object t) throws IllegalAccessException, InvocationTargetException {
        int i = 0;
        try {
            for (; i < columns.length; i++) {
                if (columns[i].method != null)
                    methodPut(i, row, t);
                else
                    fieldPut(i, row, t);
            }
        }
        catch (IllegalAccessException | InvocationTargetException ex) {
            throw ex;
        }
        catch (NumberFormatException | DateTimeException ex) {
            ListSheet.EntryColumn c = columns[i];
            String msg = "The undecorated value of cell '" + new String(int2Col(c.colIndex + 1)) + row.getRowNum() + "' is \"" + row.getString(c.colIndex) + "\"(" + row.getCellType(c.colIndex) + "), cannot cast to " + c.clazz;
            if (StringUtil.isNotEmpty(ex.getMessage())) msg = msg + ". " + ex.getMessage();
            if (detailMessageField != null) {
                detailMessageField.set(ex, msg);
                throw ex;
            } else
                throw ex instanceof DateTimeException ? new DateTimeException(msg, ex) : new NumberFormatException(msg);
        }
        catch (NullPointerException ex) {
            String msg = "Null value in cell '" + new String(int2Col(columns[i].colIndex + 1)) + row.getRowNum() + "'(" + row.getCellType(i) + ')';
            if (StringUtil.isNotEmpty(ex.getMessage())) msg = msg + ". " + ex.getMessage();
            if (detailMessageField != null) {
                detailMessageField.set(ex, msg);
                throw ex;
            } else throw new ExcelReadException(msg, ex);
        }
        catch (UncheckedTypeException ex) {
            ListSheet.EntryColumn c = columns[i];
            String msg;
            if (StringUtil.isNotEmpty(ex.getMessage())) msg ="Error occur in cell '" + new String(int2Col(c.colIndex + 1)) + row.getRowNum() + "'(" + row.getCellType(i) + "). " + ex.getMessage();
            else msg = "The undecorated value of cell '" + new String(int2Col(c.colIndex + 1)) + row.getRowNum() + "' is \"" + row.getString(c.colIndex) + "\"(" + row.getCellType(c.colIndex) + "), cannot cast to " + c.clazz;
            if (detailMessageField != null) {
                detailMessageField.set(ex, msg);
                throw ex;
            } else throw new UncheckedTypeException(msg, ex);
        }
        catch (Exception ex) {
            ListSheet.EntryColumn c = columns[i];
            String msg = "Error occur in cell '" + new String(int2Col(c.colIndex + 1)) + row.getRowNum() + "' value is \"" + row.getString(c.colIndex) + "\"(" + row.getCellType(c.colIndex) + ')';
            if (StringUtil.isNotEmpty(ex.getMessage())) msg = msg + ". " + ex.getMessage();
            if (detailMessageField != null) {
                detailMessageField.set(ex, msg);
                throw ex;
            } else throw new ExcelReadException(msg, ex);
        }
    }

    protected void fieldPut(int i, Row row, Object t) throws IllegalAccessException {
        ListSheet.EntryColumn ec = columns[i];
        int c = ec.colIndex;
        Class<?> fieldClazz = ec.clazz;
        if (fieldClazz == String.class) {
            ec.field.set(t, row.getString(c));
        }
        else if (fieldClazz == Integer.class) {
            ec.field.set(t, row.getInt(c));
        }
        else if (fieldClazz == Long.class) {
            ec.field.set(t, row.getLong(c));
        }
        else if (fieldClazz == java.util.Date.class || fieldClazz == java.sql.Date.class) {
            ec.field.set(t, row.getDate(c));
        }
        else if (fieldClazz == java.sql.Timestamp.class) {
            ec.field.set(t, row.getTimestamp(c));
        }
        else if (fieldClazz == Double.class) {
            ec.field.set(t, row.getDouble(c));
        }
        else if (fieldClazz == Float.class) {
            ec.field.set(t, row.getFloat(c));
        }
        else if (fieldClazz == Boolean.class) {
            ec.field.set(t, row.getBoolean(c));
        }
        else if (fieldClazz == BigDecimal.class) {
            ec.field.set(t, row.getDecimal(c));
        }
        else if (fieldClazz == int.class) {
            Integer v;
            ec.field.set(t, (v = row.getInt(c)) != null ? v : 0);
        }
        else if (fieldClazz == long.class) {
            Long v;
            ec.field.set(t, (v = row.getLong(c)) != null ? v : 0L);
        }
        else if (fieldClazz == double.class) {
            Double v;
            ec.field.set(t, (v = row.getDouble(c)) != null ? v : 0.0D);
        }
        else if (fieldClazz == float.class) {
            Float v;
            ec.field.set(t, (v = row.getFloat(c)) != null ? v : 0.0F);
        }
        else if (fieldClazz == boolean.class) {
            Boolean v;
            ec.field.set(t, (v = row.getBoolean(c)) != null ? v : false);
        }
        else if (fieldClazz == java.sql.Time.class) {
            ec.field.set(t, row.getTime(c));
        }
        else if (fieldClazz == LocalDateTime.class) {
            ec.field.set(t, row.getLocalDateTime(c));
        }
        else if (fieldClazz == LocalDate.class) {
            ec.field.set(t, row.getLocalDate(c));
        }
        else if (fieldClazz == LocalTime.class) {
            ec.field.set(t, row.getLocalTime(c));
        }
        else if (fieldClazz == Character.class) {
            ec.field.set(t, row.getChar(c));
        }
        else if (fieldClazz == Byte.class) {
            ec.field.set(t, row.getByte(c));
        }
        else if (fieldClazz == Short.class) {
            ec.field.set(t, row.getShort(c));
        }
        else if (fieldClazz == char.class) {
            Character v;
            ec.field.set(t, (v = row.getChar(c)) != null ? v : '\0');
        }
        else if (fieldClazz == byte.class) {
            Byte v;
            ec.field.set(t, (v = row.getByte(c)) != null ? v : 0);
        }
        else if (fieldClazz == short.class) {
            Short v;
            ec.field.set(t, (v = row.getShort(c)) != null ? v : 0);
        }
        else if (fieldClazz == RowNum.class) {
            ec.field.set(t, row.getRowNum());
        }
    }

    protected void methodPut(int i, Row row, Object t) throws IllegalAccessException, InvocationTargetException {
        ListSheet.EntryColumn ec = columns[i];
        int c = ec.colIndex;
        Class<?> fieldClazz = ec.clazz;
        if (fieldClazz == String.class) {
            ec.method.invoke(t, row.getString(c));
        }
        else if (fieldClazz == Integer.class) {
            ec.method.invoke(t, row.getInt(c));
        }
        else if (fieldClazz == Long.class) {
            ec.method.invoke(t, row.getLong(c));
        }
        else if (fieldClazz == java.util.Date.class || fieldClazz == java.sql.Date.class) {
            ec.method.invoke(t, row.getDate(c));
        }
        else if (fieldClazz == java.sql.Timestamp.class) {
            ec.method.invoke(t, row.getTimestamp(c));
        }
        else if (fieldClazz == Double.class) {
            ec.method.invoke(t, row.getDouble(c));
        }
        else if (fieldClazz == Float.class) {
            ec.method.invoke(t, row.getFloat(c));
        }
        else if (fieldClazz == Boolean.class) {
            ec.method.invoke(t, row.getBoolean(c));
        }
        else if (fieldClazz == BigDecimal.class) {
            ec.method.invoke(t, row.getDecimal(c));
        }
        else if (fieldClazz == int.class) {
            Integer v;
            ec.method.invoke(t, (v = row.getInt(c)) != null ? v : 0);
        }
        else if (fieldClazz == long.class) {
            Long v;
            ec.method.invoke(t, (v = row.getLong(c)) != null ? v : 0);
        }
        else if (fieldClazz == double.class) {
            Double v;
            ec.method.invoke(t, (v = row.getDouble(c)) != null ? v : 0.0D);
        }
        else if (fieldClazz == float.class) {
            Float v;
            ec.method.invoke(t, (v = row.getFloat(c)) != null ? v : 0.0F);
        }
        else if (fieldClazz == boolean.class) {
            Boolean v;
            ec.method.invoke(t, (v = row.getBoolean(c)) != null ? v : false);
        }
        else if (fieldClazz == java.sql.Time.class) {
            ec.method.invoke(t, row.getTime(c));
        }
        else if (fieldClazz == LocalDateTime.class) {
            ec.method.invoke(t, row.getLocalDateTime(c));
        }
        else if (fieldClazz == LocalDate.class) {
            ec.method.invoke(t, row.getLocalDate(c));
        }
        else if (fieldClazz == LocalTime.class) {
            ec.method.invoke(t, row.getLocalTime(c));
        }
        else if (fieldClazz == Character.class) {
            ec.method.invoke(t, row.getChar(c));
        }
        else if (fieldClazz == Byte.class) {
            ec.method.invoke(t, row.getByte(c));
        }
        else if (fieldClazz == Short.class) {
            ec.method.invoke(t, row.getShort(c));
        }
        else if (fieldClazz == char.class) {
            Character v;
            ec.method.invoke(t, (v = row.getChar(c)) != null ? v : '\0');
        }
        else if (fieldClazz == byte.class) {
            Byte v;
            ec.method.invoke(t, (v = row.getByte(c)) != null ? v : 0);
        }
        else if (fieldClazz == short.class) {
            Short v;
            ec.method.invoke(t, (v = row.getShort(c)) != null ? v : 0);
        }
        else if (fieldClazz == RowNum.class) {
            ec.method.invoke(t, row.getRowNum());
        }
    }

    /**
     * Ignore some columns, override this method to add custom filtering
     *
     * @param ao {@code Method} or {@code Field}
     * @return true if ignore current column
     */
    protected boolean ignoreColumn(AccessibleObject ao) {
        return ao.getAnnotation(IgnoreImport.class) != null;
    }

    /**
     * Attach some others column
     *
     * @param clazz Target class
     * @return a column and Write method mapping
     */
    protected Map<String, Method> attachOtherColumn(Class<?> clazz) {
        Method[] writeMethods;
        try {
            writeMethods = (option & 1) == 1 ? listDeclaredMethods(clazz)
                : listDeclaredMethods(clazz, method -> method.getAnnotation(ExcelColumn.class) != null || method.getAnnotation(RowNum.class) != null);
        } catch (IntrospectionException e) {
            LOGGER.warn("Get [" + clazz + "] read declared failed.", e);
            return Collections.emptyMap();
        }

        return Arrays.stream(writeMethods).filter(m -> m.getParameterCount() == 1).collect(Collectors.toMap(a -> { String k = a.getName(); return k.startsWith("set") ? lowFirstKey(k.substring(3)) : k; }, a -> a, (a, b) -> b));
    }

    /**
     * Create column from {@link ExcelColumn} annotation
     * <p>
     * Override the method to extend custom comments
     *
     * @param ao {@link AccessibleObject} witch defined the {@code ExcelColumn} annotation
     * @return the Worksheet's {@link ListSheet.EntryColumn} information
     */
    protected ListSheet.EntryColumn createColumn(AccessibleObject ao) {
        // Filter all ignore column
        if (ignoreColumn(ao)) return null;

        ao.setAccessible(true);
        // Support multi header columns
        ExcelColumns cs = ao.getAnnotation(ExcelColumns.class);
        if (cs != null) {
            ExcelColumn[] ecs = cs.value();
            ListSheet.EntryColumn root = null;
            for (int i = Math.max(0, ecs.length - headRows); i < ecs.length; i++) {
                ListSheet.EntryColumn column = createColumnByAnnotation(ecs[i]);
                if (root == null) {
                    root = column;
                } else {
                    root.addSubColumn(column);
                }
            }
            return root;
        }
        // Single header column
        ExcelColumn ec = ao.getAnnotation(ExcelColumn.class);
        if (ec != null) return createColumnByAnnotation(ec);
        // Row Num
        RowNum rowNum = ao.getAnnotation(RowNum.class);
        if (rowNum != null) return createColumnByAnnotation(rowNum);

        ListSheet.EntryColumn column = null;
        // Direct match method or field name
        if ((option & 1) == 1) {
            if (Field.class.isAssignableFrom(ao.getClass())) {
                Field f = (Field) ao;
                column = new ListSheet.EntryColumn(f.getName());
                column.field = f;
                column.clazz = f.getType();
            }
            else if (Method.class.isAssignableFrom(ao.getClass())) {
                Method m = (Method) ao;
                String k = m.getName();
                column = new ListSheet.EntryColumn(k.startsWith("set") ? lowFirstKey(k.substring(3)) : k);
                column.method = m;
                if (m.getParameterCount() == 1) column.clazz = m.getParameterTypes()[0];
            }
        }
        return column;
    }

    /**
     * Create column by {@code ExcelColumn} annotation
     *
     * @param anno an java annotation
     * @return {@link Column} or null if annotation is null
     */
    protected ListSheet.EntryColumn createColumnByAnnotation(Annotation anno) {
        ListSheet.EntryColumn column = null;
        if (anno instanceof ExcelColumn) {
            ExcelColumn ec = (ExcelColumn) anno;
            column = new ListSheet.EntryColumn(ec.value());
            column.setColIndex(ec.colIndex());
            // Hidden Column
            if (ec.hide()) column.hide();
        } else if (anno instanceof RowNum) {
            column = new ListSheet.EntryColumn(EMPTY, RowNum.class);
        }
        return column;
    }

    /**
     * Copy column name on merge cells
     *
     * @param mergeCells merged cell in header rows
     * @param rows the header rows
     */
    protected void mergeCellsIfNull(List<Dimension> mergeCells, Row[] rows) {
        if (mergeCells == null) return;
        mergeCells = mergeCells.stream().filter(d -> d.getWidth() > 1).collect(Collectors.toList());
        if (mergeCells.isEmpty()) return;

        Map<Long, Dimension> map = mergeCells.stream().collect(Collectors.toMap(a -> ((long) a.firstRow) << 16 | a.firstColumn, a -> a, (a, b) -> a));

        for (Row row : rows) {
            for (int i = row.fc; i < row.lc; ) {
                Cell cell = row.cells[i];
                Dimension d = map.get(((long) row.getRowNum()) << 16 | cell.i);
                if (d != null) {
                    String v = row.getString(cell);
                    if (d.lastColumn > row.lc) {
                        row.cells = row.copyCells(d.lastColumn);
                        row.lc = d.lastColumn;
                    }
                    for (int j = d.firstColumn + 1; j <= d.lastColumn; j++) row.cells[++i].setSv(v);
                } else i++;
            }
        }
    }

    public HeaderRow setOptions(int option) {
        this.option = option;
        if (mapping != null && (option & 6) > 0) {
            Map<String, Integer> m = new HashMap<>(mapping.size());
            for (Map.Entry<String, Integer> entry : mapping.entrySet()) {
                m.put(makeKey(entry.getKey()), entry.getValue());
            }
            this.mapping = m;
        }
        return this;
    }

    protected String makeKey(String key) {
        if ((option & 4) == 4) key = toCamelCase(key);
        if ((option & 2) == 2) key = key.toLowerCase();
        return key;
    }
}

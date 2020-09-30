/*
 * Copyright (c) 2017-2018, guanquan.wang@yandex.com All Rights Reserved.
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

package org.ttzero.excel.entity;

import org.ttzero.excel.annotation.ExcelColumn;
import org.ttzero.excel.annotation.HeaderComment;
import org.ttzero.excel.reader.Cell;
import org.ttzero.excel.annotation.IgnoreExport;

import java.beans.IntrospectionException;
import java.beans.Introspector;
import java.beans.PropertyDescriptor;
import java.io.IOException;
import java.lang.reflect.AccessibleObject;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.lang.reflect.ParameterizedType;
import java.util.ArrayList;
import java.util.Collection;
import java.util.HashMap;
import java.util.List;
import java.util.Map;


import static org.ttzero.excel.util.ReflectUtil.listDeclaredFields;
import static org.ttzero.excel.util.ReflectUtil.listReadMethods;
import static org.ttzero.excel.util.StringUtil.EMPTY;
import static org.ttzero.excel.util.StringUtil.isNotEmpty;
import static org.ttzero.excel.util.StringUtil.isEmpty;

/**
 * List is the most important data source, you can pass all
 * the data at a time, or customize the worksheet to extends
 * the {@code ListSheet}, and then override the {@link #more}
 * method to achieve segmented loading of data. The {@link #more}
 * method returns NULL or an empty array to complete the current
 * worksheet write
 *
 * @see ListMapSheet
 *
 * @author guanquan.wang at 2018-01-26 14:48
 */
public class ListSheet<T> extends Sheet {
    protected List<T> data;
    private Field[] fields;
    private Method[] methods;
    protected int start, end;
    protected boolean eof;
    private int size;

    /**
     * Constructor worksheet
     */
    public ListSheet() {
        super();
    }

    /**
     * Constructor worksheet
     *
     * @param name the worksheet name
     */
    public ListSheet(String name) {
        super(name);
    }

    /**
     * Constructor worksheet
     *
     * @param name    the worksheet name
     * @param columns the header info
     */
    public ListSheet(String name, final Column... columns) {
        super(name, columns);
    }

    /**
     * Constructor worksheet
     *
     * @param name      the worksheet name
     * @param waterMark the water mark
     * @param columns   the header info
     */
    public ListSheet(String name, WaterMark waterMark, final Column... columns) {
        super(name, waterMark, columns);
    }

    /**
     * Constructor worksheet
     *
     * @param data the worksheet's body data
     */
    public ListSheet(List<T> data) {
        this(null, data);
    }

    /**
     * Constructor worksheet
     *
     * @param name the worksheet name
     * @param data the worksheet's body data
     */
    public ListSheet(String name, List<T> data) {
        super(name);
        setData(data);
    }

    /**
     * Constructor worksheet
     *
     * @param data    the worksheet's body data
     * @param columns the header info
     */
    public ListSheet(List<T> data, final Column... columns) {
        this(null, data, columns);
    }

    /**
     * Constructor worksheet
     *
     * @param name    the worksheet name
     * @param data    the worksheet's body data
     * @param columns the header info
     */
    public ListSheet(String name, List<T> data, final Column... columns) {
        this(name, data, null, columns);
    }

    /**
     * Constructor worksheet
     *
     * @param data      the worksheet's body data
     * @param waterMark the water mark
     * @param columns   the header info
     */
    public ListSheet(List<T> data, WaterMark waterMark, final Column... columns) {
        this(null, data, waterMark, columns);
    }

    /**
     * Constructor worksheet
     *
     * @param name      the worksheet name
     * @param data      the worksheet's body data
     * @param waterMark the water mark
     * @param columns   the header info
     */
    public ListSheet(String name, List<T> data, WaterMark waterMark, final Column... columns) {
        super(name, waterMark, columns);
        setData(data);
    }

    /**
     * Setting the worksheet data
     *
     * @param data the body data
     * @return worksheet
     */
    public ListSheet<T> setData(final List<T> data) {
        this.data = data;
        if (!headerReady && workbook != null) {
            getHeaderColumns();
        }
        // Has data and worksheet can write
        // Paging in advance
        if (data != null && sheetWriter != null) {
            paging();
        }
        return this;
    }

    /**
     * Returns the first not null object
     *
     * @return the object
     */
    protected T getFirst() {
        if (data == null || data.isEmpty()) return null;
        T first = data.get(start);
        if (first != null) return first;
        int i = start + 1;
        do {
            first = data.get(i++);
        } while (first == null);
        return first;
    }

    /**
     * Release resources
     *
     * @throws IOException if I/O error occur
     */
    @Override
    public void close() throws IOException {
        // Maybe there has more data
        if (!eof && rows >= sheetWriter.getRowLimit() - 1) {
            List<T> list = more();
            if (list != null && !list.isEmpty()) {
                compact();
                data.addAll(list);
                @SuppressWarnings("unchecked")
                ListSheet<T> copy = getClass().cast(clone());
                copy.start = 0;
                copy.end = list.size();
                workbook.insertSheet(id, copy);
                // Do not close current worksheet
                shouldClose = false;
            }
        }
        if (shouldClose && data != null) {
            // Some Collection not support #remove
//            data.clear();
            data = null;
        }
        super.close();
    }

    /**
     * Reset the row-block data
     */
    @Override
    protected void resetBlockData() {
        if (!eof && left() < getRowBlockSize()) {
            append();
        }

        // Find the end index of row-block
        int end = getEndIndex();
        int len = columns.length;
        try {
            for (; start < end; rows++, start++) {
                Row row = rowBlock.next();
                row.index = rows;
                Cell[] cells = row.realloc(len);
                T o = data.get(start);
                for (int i = 0; i < len; i++) {
                    // clear cells
                    Cell cell = cells[i];
                    cell.clear();

                    Object e;
                    if (columns[i].isIgnoreValue())
                        e = null;
                    else if (methods[i] != null)
                        e = methods[i].invoke(o);
                    else
                        e = fields[i].get(o);

                    cellValueAndStyle.reset(rows, cell, e, columns[i]);
                }
            }
        } catch (IllegalAccessException | InvocationTargetException e) {
            throw new ExcelWriteException(e);
        }
    }

    /**
     * Call this method to get more data when the data length
     * less than the row-block size until there is no more data
     * or more than the row limit
     */
    protected void append() {
        int rbs = getRowBlockSize();
        for (; ; ) {
            List<T> list = more();
            // No more data
            if (list == null || list.isEmpty()) {
                eof = shouldClose = true;
                break;
            }
            // The first getting
            if (data == null) {
                setData(list);

                if (list.size() < rbs) continue;
                else break;
            }
            compact();
            data.addAll(list);
            start = 0;
            end = data.size();
            // Split worksheet
            if (end >= rbs) {
                paging();
                break;
            }
        }
    }

    private void compact() {
        // Copy the remaining data to a temporary array
        int size = left();
        if (start > 0 && size > 0) {
            // append and resize
            List<T> last = new ArrayList<>(size);
            last.addAll(data.subList(start, end));
            data.clear();
            data.addAll(last);
        } else if (start > 0) data.clear();
    }

    // Returns the reflect <T> type
    private Class<?> getTClass() {
        Class<?> clazz;
        if (getClass().getGenericSuperclass() instanceof ParameterizedType) {
            @SuppressWarnings({"unchecked", "retype"})
            Class<?> tClass = (Class<T>) ((ParameterizedType) getClass()
                .getGenericSuperclass()).getActualTypeArguments()[0];
            clazz = tClass;
        } else {
            T o = getFirst();
            if (o == null) return null;
            clazz = o.getClass();
        }
        return clazz;
    }

    /**
     * Get the first object of the object array witch is not NULL,
     * reflect all declared fields, and then do the following steps
     * <p>
     * step 1. If the method has {@link ExcelColumn} annotation, the value of
     * this annotation is used as the column name.
     * <p>
     * step 2. If the {@link ExcelColumn} annotation has no value or empty value,
     * the field name is used as the column name.
     * <p>
     * step 3. If the field has {@link ExcelColumn} annotation, the value of
     * this annotation is used as the column name.
     * <p>
     * step 4. Skip this Field if it has a {@link IgnoreExport} annotation,
     * or the field which has not {@link ExcelColumn} annotation.
     * <p>
     * The column order is the same as the order in declared fields.
     *
     * @return the column array length
     */
    private int init() {
        Class<?> clazz = getTClass();
        if (clazz == null) return 0;

        Map<String, Method> tmp = new HashMap<>();
        try {
            PropertyDescriptor[] propertyDescriptors = Introspector.getBeanInfo(clazz)
                    .getPropertyDescriptors();
            for (PropertyDescriptor pd : propertyDescriptors) {
                Method method = pd.getReadMethod();
                if (method == null) continue;
                tmp.put(pd.getName(), method);
            }
        } catch (IntrospectionException e) {
            LOGGER.warn("Get class {} methods failed.", clazz);
        }

        Field[] declaredFields = listDeclaredFields(clazz);

        if (!hasHeaderColumns()) {
            // Get ExcelColumn annotation method
            methods = new Method[declaredFields.length + tmp.size()];
            List<Column> list = new ArrayList<>(declaredFields.length);

            int i = 0;
            for (; i < declaredFields.length; i++) {
                Field field = declaredFields[i];
                field.setAccessible(true);
                String gs = field.getName();

                // Ignore annotation on read method
                Method method = tmp.get(gs);
                if (method != null) {
                    if (method.getAnnotation(IgnoreExport.class) != null) {
                        declaredFields[i] = null;
                        continue;
                    }

                    Column column = createColumn(method);
                    if (column != null) {
                        methods[i] = method;
                        column.clazz = method.getReturnType();
                        column.key = gs;
                        if (isEmpty(column.name)) {
                            column.name = gs;
                        }
                        list.add(column);
                        continue;
                    }
                }

                Column column = createColumn(field);
                if (column != null) {
                    list.add(column);
                    column.key = gs;
                    if (isEmpty(column.name)) {
                        column.name = gs;
                    }
                    if (method != null) {
                        column.clazz = method.getReturnType();
                        methods[i] = method;
                    } else column.clazz = field.getType();
                    continue;
                }

                // Ignore others
                declaredFields[i] = null;
            }

            // Collect the method which has ExcelColumn annotation
            Method[] readMethods = null;
            try {
                Collection<Method> values = tmp.values();
                readMethods = listReadMethods(clazz, method -> method.getAnnotation(ExcelColumn.class) != null
                        && !values.contains(method));
            } catch (IntrospectionException e) {
                // Ignore
            }

            if (readMethods != null) {
                for (Method method : readMethods) {
                    Column column = createColumn(method);
                    if (column != null) {
                        methods[i++] = method;
                        list.add(column);
                        column.clazz = method.getReturnType();
                        column.key = method.getName();
                        if (isEmpty(column.name)) {
                            column.name = method.getName();
                        }
                    }
                }
            }

            // No column to write
            if (list.isEmpty()) {
                headerReady = eof = shouldClose = true;
                this.end = 0;
                LOGGER.warn("Class [{}] do not contains properties to export.", clazz);
                return 0;
            }
            columns = new Column[list.size()];
            list.toArray(columns);
            for (i = 0; i < columns.length; i++) {
                columns[i].styles = workbook.getStyles();
            }

            // Clean
            i = 0;
            fields = new Field[columns.length];
            for (int j = 0; j < declaredFields.length; j++) {
                if (declaredFields[j] != null) {
                    declaredFields[j].setAccessible(true);
                    fields[i] = declaredFields[j];
                    methods[i] = methods[j];
                    i++;
                }
            }
            if (declaredFields.length < methods.length) {
                System.arraycopy(methods, declaredFields.length, methods, i, methods.length - declaredFields.length);
                i += methods.length - declaredFields.length;
            }
            return i;
        } else {
            fields = new Field[columns.length];
            methods = new Method[columns.length];
            for (int i = 0; i < columns.length; i++) {
                Column hc = columns[i];
                methods[i] = tmp.get(hc.key);
                if (methods[i] != null) methods[i].setAccessible(true);

                for (Field field : declaredFields) {
                    if (hc.key.equals(field.getName())) {
                        field.setAccessible(true);
                        fields[i] = field;
                        break;
                    }
                }

                if (methods[i] == null && fields[i] == null) {
                    LOGGER.warn("Column [" + hc.getName() + "(" + hc.key + ")"
                            + "] not declare in class " + clazz);
                    hc.ignoreValue();
                } else if (hc.getClazz() == null) {
                    hc.setClazz(methods[i] != null ? methods[i].getReturnType() : fields[i].getType());
                }
            }
            return columns.length;
        }
    }

    private Column createColumn(AccessibleObject ao) {
        if (ao.getAnnotation(IgnoreExport.class) != null) return null;
        ao.setAccessible(true);
        ExcelColumn ec = ao.getAnnotation(ExcelColumn.class);
        if (ec != null) {
            Column column = new Column(ec.value(), EMPTY, ec.share());
            // Comment
            column.headerComment = createComment(ao.getAnnotation(HeaderComment.class), ec.comment());
            // Number format
            if (isNotEmpty(ec.format())) {
                column.setNumFmt(ec.format());
            }
            return column;
        }
        return null;
    }

    private Comment createComment(HeaderComment precedence, HeaderComment normal) {
        HeaderComment comment = precedence != null ? precedence : normal;
        if (comment != null && (isNotEmpty(comment.value()) || isNotEmpty(comment.title()))) {
            return new Comment(comment.title(), comment.value());
        }
        return null;
    }

    /**
     * Returns the header column info
     *
     * @return array of column
     */
    @Override
    public Column[] getHeaderColumns() {
        if (!headerReady) {
//            if (!hasHeaderColumns()) {
//                columns = new Column[0];
//            }
            // create header columns
            int size = init();
            if (size <= 0) {
                columns = new Column[0];
            } else {
                // Check the header column limit
                checkColumnLimit();
                headerReady = true;
            }
        }
        return columns;
    }

    /**
     * Returns the end index of row-block
     *
     * @return the end index
     */
    protected int getEndIndex() {
        int blockSize = getRowBlockSize(), rowLimit = sheetWriter.getRowLimit() - 1;
        if (rows + blockSize > rowLimit) {
            blockSize = rowLimit - rows;
        }
        int end = start + blockSize;
        return Math.min(end, this.end);
    }

    /**
     * Returns total rows in this worksheet
     *
     * @return -1 if unknown or uncertain
     */
    @Override
    public int size() {
        return !shouldClose ? size : -1;
    }

    /**
     * Returns left data in array to be write
     *
     * @return the left number
     */
    protected int left() {
        return end - start;
    }

    /**
     * Split worksheet data
     */
    @Override
    protected void paging() {
        int len = dataSize(), limit = sheetWriter.getRowLimit() - 1;
        // paging
        if (len + rows > limit) {
            // Reset current index
            end = limit - rows + start;
            shouldClose = false;
            eof = true;
            size = limit;

            int n = id;
            for (int i = end; i < len; ) {
                @SuppressWarnings("unchecked")
                ListSheet<T> copy = getClass().cast(clone());
                copy.start = i;
                copy.end = (i = Math.min(i + limit, len));
                copy.size = copy.end - copy.start;
                copy.eof = copy.size == limit;
                workbook.insertSheet(n++, copy);
            }
            // Close on the last copy worksheet
            workbook.getSheetAt(n - 1).shouldClose = true;
        } else {
            end = len;
            size += len;
        }
    }

    /**
     * Returns total data size before split
     *
     * @return the total size
     */
    public int dataSize() {
        return data != null ? data.size() : 0;
    }

    /**
     * This method is used for the worksheet to get the data.
     * This is a data source independent method. You can get data
     * from any data source. Since this method is stateless, you
     * should manage paging or other information in your custom Sheet.
     * <p>
     * The more data you get each time, the faster write speed. You
     * should minimize the database query or network request, but the
     * excessive data will put pressure on the memory. Please balance
     * this value between the speed and memory. You can refer to {@code 2^8 ~ 2^10}
     * <p>
     * This method is blocked
     *
     * @return The data output to the worksheet, if a null or empty
     * ArrayList returned, mean that the current worksheet is finished written.
     */
    protected List<T> more() {
        return null;
    }
}

/*
 * Copyright (c) 2019, guanquan.wang@yandex.com All Rights Reserved.
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

package cn.ttzero.excel.entity;

import cn.ttzero.excel.reader.Cell;
import cn.ttzero.excel.util.StringUtil;
import cn.ttzero.excel.annotation.DisplayName;
import cn.ttzero.excel.annotation.NotExport;

import java.io.IOException;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;

/**
 * Created by guanquan.wang at 2018-01-26 14:48
 */
public class ListObjectSheet<T> extends ListSheet {
    private List<T> data;
    private Field[] fields;

    /**
     * Constructor worksheet
     */
    public ListObjectSheet() {
        super();
    }

    /**
     * Constructor worksheet
     * @param name the worksheet name
     */
    public ListObjectSheet(String name) {
        super(name);
    }

    /**
     * Constructor worksheet
     * @param name the worksheet name
     */
    public ListObjectSheet(String name, Column[] columns) {
        super(name, columns);
    }

    /**
     * Constructor worksheet
     * @param name the worksheet name
     */
    public ListObjectSheet(String name, WaterMark waterMark, Column[] columns) {
        super(name, waterMark, columns);
    }

    public ListObjectSheet<T> setData(final List<T> data) {
        this.data = data;
        if (data != null) {
            end = data.size();
        }
        return this;
    }

    /**
     * Release resources
     * @throws IOException if io error occur
     */
    @Override
    public void close() throws IOException {
        if (shouldClose) {
            data.clear();
            data = null;
        }
        super.close();
    }

    /**
     * Reset the row-block data
     */
    @Override
    protected void resetBlockData() {
        for (int rbs = getRowBlockSize(), size; (size = size()) < rbs; ) {
            List<T> list = more();
            if (list == null || list.isEmpty()) break;
            if (data == null) {
                data = new ArrayList<>(rbs);
            }
            if (start > 0 && size > 0) {
                // append and resize
                List<T> last = new ArrayList<>(size);
                last.addAll(data.subList(start, end));
                data.clear();
                data.addAll(last);
            } else data.clear();
            data.addAll(list);
            start = 0;
            end = data.size();
        }
        if (!headerReady) {
            getHeaderColumns();
        }
        // Find the end index of row-block
        int end = getEndIndex();
        int len = columns.length;
        try {
            for (; start < end; rows++, start++) {
                Row row = rowBlock.next();
                row.index = rows;
                Field field;
                Cell[] cells = row.realloc(len);
                for (int i = 0; i < len; i++) {
                    field = fields[i];
                    // clear cells
                    Cell cell = cells[i];
                    cell.clear();

                    Object e = field.get(data.get(start));
                    // blank cell
                    if (e == null) {
                        cell.setBlank();
                        continue;
                    }

                    setCellValueAndStyle(cell, e, columns[i]);
                }
            }
        } catch (IllegalAccessException e) {
            throw new ExcelWriteException(e);
        }
    }

    private static final String[] exclude = {"serialVersionUID", "this$0"};

    private Field[] init() {
        Object o = workbook.getFirst(data);
        if (o == null) return null;
        if (columns == null || columns.length == 0) {
            Field[] fields = o.getClass().getDeclaredFields();
            List<Column> list = new ArrayList<>(fields.length);
            for (int i = 0; i < fields.length; i++) {
                Field field = fields[i];
                String gs = field.toGenericString();
                NotExport notExport = field.getAnnotation(NotExport.class);
                if (notExport != null || StringUtil.indexOf(exclude, gs.substring(gs.lastIndexOf('.') + 1)) >= 0) {
                    fields[i] = null;
                    continue;
                }
                DisplayName dn = field.getAnnotation(DisplayName.class);
                if (dn != null && StringUtil.isNotEmpty(dn.value())) {
                    list.add(new Column(dn.value(), field.getName(), field.getType()).setShare(dn.share()));
                } else {
                    list.add(new Column(field.getName(), field.getName(), field.getType()).setShare(dn != null && dn.share()));
                }
            }
            columns = new Column[list.size()];
            list.toArray(columns);
            for (int i = 0; i < columns.length; i++) {
                columns[i].styles = workbook.getStyles();
            }
            // clear not export fields
            for (int len = fields.length, n = len - 1; n >= 0; n--) {
                if (fields[n] != null) {
                    fields[n].setAccessible(true);
                    continue;
                }
                if (n < len - 1) {
                    System.arraycopy(fields, n + 1, fields, n, len - n - 1);
                }
                len--;
            }
            return fields;
        } else {
            Field[] fields = new Field[columns.length];
            Class<?> clazz = o.getClass();
            for (int i = 0; i < columns.length; i++) {
                Column hc = columns[i];
                try {
                    fields[i] = clazz.getDeclaredField(hc.key);
                    fields[i].setAccessible(true);
                    if (hc.getClazz() == null) {
                        hc.setClazz(fields[i].getType());
//                        DisplayName dn = field.getAnnotation(DisplayName.class);
//                        if (dn != null) {
//                            hc.setShare(hc.isShare() || dn.share());
//                            if (StringUtil.isEmpty(hc.getName())
//                                    && StringUtil.isNotEmpty(dn.value())) {
//                                hc.setName(dn.value());
//                            }
//                        }
                    }
                } catch (NoSuchFieldException e) {
                    throw new ExcelWriteException("Column " + hc.getName() + " not declare in class " + clazz);
                }
            }
            return fields;
        }

    }

    /**
     * Returns the header column info
     * @return array of column
     */
    @Override
    public Column[] getHeaderColumns() {
        if (!headerReady) {
            if (data == null || data.isEmpty()) {
                columns = new Column[0];
            }
            // create header columns
            fields = init();
            if (fields == null || fields.length == 0 || fields[0] == null) {
                columns = new Column[0];
            } else {
                headerReady = true;
            }
        }
        return columns;
    }

    /**
     * Paging worksheet
     * @return a copy worksheet
     */
    @Override
    public ListObjectSheet<T> copy() {
        ListObjectSheet<T> sheet = new ListObjectSheet<>(name, columns);
        sheet.workbook = workbook;
        sheet.data = data;
        sheet.autoSize = autoSize;
        sheet.autoOdd = autoOdd;
        sheet.oddFill = oddFill;
        sheet.relManager = relManager.clone();
        sheet.sheetWriter = sheetWriter.copy(sheet);
        sheet.waterMark = waterMark;
        sheet.copySheet = true;
        return sheet;
    }

    /**
     * Returns total data size before split
     * @return the total size
     */
    @Override
    public int dataSize() {
        return data != null ? data.size() : 0;
    }

    /**
     * Get more row data if hasMore returns true
     * @return the row data
     */
    @Override
    protected List<T> more() {
        return null;
    }
}

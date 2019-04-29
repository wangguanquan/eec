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
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;

import static cn.ttzero.excel.manager.Const.ROW_BLOCK_SIZE;

/**
 * Created by guanquan.wang at 2018-01-26 14:48
 */
public class ListObjectSheet<T> extends Sheet {
    private List<T> data;
    private Field[] fields;
    private int start, end;

    public ListObjectSheet(Workbook workbook) {
        super(workbook);
    }

    public ListObjectSheet(Workbook workbook, String name, Column[] columns) {
        super(workbook, name, columns);
    }

    public ListObjectSheet(Workbook workbook, String name, WaterMark waterMark, Column[] columns) {
        super(workbook, name, waterMark, columns);
    }


    @Override
    public void close() throws IOException {
        if (shouldClose) {
            data.clear();
            data = null;
        }
        super.close();
    }

    public ListObjectSheet<T> setData(final List<T> data) {
        this.data = data;
        return this;
    }

    /**
     * write worksheet data to path
     * @param path the storage path
     * @throws IOException write error
     * @throws ExcelWriteException others
     */
    public void writeTo(Path path) throws IOException, ExcelWriteException {
        if (sheetWriter == null) {
            throw new ExcelWriteException("Worksheet writer is not instanced.");
        }
        if (!copySheet) {
            paging();
        }
        rowBlock = new RowBlock();
        sheetWriter.write(path);
    }

    @Override
    public RowBlock nextBlock() {
        // clear first
        rowBlock.clear();

        try {
            loopData();
        } catch (IllegalAccessException e) {
            throw new ExcelWriteException(e);
        }

        return rowBlock.flip();
    }

    private void loopData() throws IllegalAccessException {
        // Find the end index of row-block
        int end = getEndIndex();
        int len = columns.length;
        for ( ; start < end; rows++, start++) {
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

                setCellValue(cell, e, columns[i]);
            }
        }
    }

    private int getEndIndex() {
        int end = start + rows + ROW_BLOCK_SIZE;
        return end <= this.end ? end : this.end;
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
            }
            headerReady = true;
        }
        return columns;
    }

    /**
     * Returns total rows in this worksheet
     * @return -1 if unknown
     */
    @Override
    public int size() {
        return end - start;
    }

    protected void paging() {
        int len = data.size(), limit = sheetWriter.getRowLimit() - 1;
        workbook.what("Total size: " + len);
        // paging
        if (len > limit) {
            int page = len / limit;
            if (len % limit > 0) {
                page++;
            }
            // Insert sub-sheet
            for (int i = 1, index = id, last = page - 1, n; i < page; i++) {
                ListObjectSheet<T> sheet = copy();
                sheet.name = name + " (" + i + ")";
                sheet.start = i * limit;
                sheet.end = (n = (i + 1) * limit) < len ? n : len;
                sheet.shouldClose = i == last;
                workbook.insertSheet(index++, sheet);
            }
            // Reset current index
            start = 0;
            end = limit;
            shouldClose = false;
        } else {
            start = 0;
            end = len;
        }
    }

    public ListObjectSheet<T> copy() {
        ListObjectSheet<T> sheet = new ListObjectSheet<>(workbook, name, columns);
        sheet.data = data;
        sheet.autoSize = autoSize;
        sheet.autoOdd = autoOdd;
        sheet.oddFill = oddFill;
        sheet.sheetWriter = sheetWriter.copy(sheet);
        sheet.copySheet = true;
        return sheet;
    }
}

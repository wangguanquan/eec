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

import java.io.IOException;
import java.nio.file.Path;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import static cn.ttzero.excel.manager.Const.ROW_BLOCK_SIZE;

/**
 * Created by guanquan.wang at 2018-01-26 14:46
 */
public class ListMapSheet extends Sheet {
    private List<Map<String, ?>> data;
    private int start, end;

    public ListMapSheet(Workbook workbook) {
        super(workbook);
    }

    public ListMapSheet(Workbook workbook, String name, Column[] columns) {
        super(workbook, name, columns);
    }

    public ListMapSheet(Workbook workbook, String name, WaterMark waterMark, Column[] columns) {
        super(workbook, name, waterMark, columns);
    }

    /**
     * Returns the header column info
     *
     * @return array of column
     */
    @Override
    public Column[] getHeaderColumns() {
        if (headerReady) return columns;
        @SuppressWarnings("unchecked")
        Map<String, ?> first = (Map<String, ?>) workbook.getFirst(data);
        // No data
        if (first == null) {
            if (columns == null) {
                columns = new Column[0];
            }
        }
        else if (columns.length == 0) {
            int size = first.size(), i = 0;
            columns = new Column[size];
            for (Iterator<? extends Map.Entry<String, ?>> it = first.entrySet().iterator(); it.hasNext(); ) {
                Map.Entry<String, ?> entry = it.next();
                columns[i++] = new Column(entry.getKey(), entry.getKey(), entry.getValue().getClass());
            }
        }
        else {
            for (int i = 0; i < columns.length; i++) {
                Column hc = columns[i];
                if (StringUtil.isEmpty(hc.key)) {
                    throw new ExcelWriteException(getClass() + " 类别必须指定map的key。");
                }
                if (hc.getClazz() == null) {
                    hc.setClazz(first.get(hc.key).getClass());
                }
            }
        }
        for (Column hc : columns) {
            hc.styles = workbook.getStyles();
        }
        headerReady = true;
        return columns;
    }

    @Override
    public void close() throws IOException {
        if (shouldClose) {
            data.clear();
            data = null;
        }
        super.close();
    }

    public ListMapSheet setData(final List<Map<String, ?>> data) {
        this.data = data;
        return this;
    }

    /**
     * Returns a row-block. The row-block is content by 32 rows
     *
     * @return a row-block
     */
    @Override
    public RowBlock nextBlock() {
        // clear first
        rowBlock.clear();

        loopData();

        return rowBlock.flip();
    }

    private void loopData() {
        int end = getEndIndex();
        int len = columns.length;
        for (; start < end; rows++, start++) {
            Row row = rowBlock.next();
            row.index = rows;
            Cell[] cells = row.realloc(len);
            for (int i = 0; i < len; i++) {
                Column hc = columns[i];
                Object e = data.get(start).get(hc.key);
                // clear cells
                Cell cell = cells[i];
                cell.clear();

                // blank cell
                if (e == null) {
                    cell.setBlank();
                    continue;
                }

                setCellValue(cell, e, hc);
            }
        }
    }

    private int getEndIndex() {
        int end = start + rows + ROW_BLOCK_SIZE;
        return end <= this.end ? end : this.end;
    }

    /**
     * Returns total rows in this worksheet
     *
     * @return -1 if unknown
     */
    public int size() {
        return end - start;
    }

    /**
     * Split worksheet data
     */
    @Override
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
                ListMapSheet sheet = copy();
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

    public ListMapSheet copy() {
        ListMapSheet sheet = new ListMapSheet(workbook, name, columns);
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
}

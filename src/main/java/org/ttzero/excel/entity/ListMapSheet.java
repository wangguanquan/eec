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

import org.ttzero.excel.manager.Const;
import org.ttzero.excel.reader.Cell;
import org.ttzero.excel.reader.Dimension;
import org.ttzero.excel.reader.Grid;
import org.ttzero.excel.reader.GridFactory;
import org.ttzero.excel.util.StringUtil;

import java.util.Arrays;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import static org.ttzero.excel.util.StringUtil.isEmpty;

/**
 * ListMapSheet is a subclass of {@link ListSheet}, the difference is
 * in the way the data is taken
 *
 * @see ListSheet
 *
 * @author guanquan.wang at 2018-01-26 14:46
 */
public class ListMapSheet extends ListSheet<Map<String, ?>> {

    /**
     * Constructor worksheet
     */
    public ListMapSheet() {
        super();
    }

    /**
     * Constructor worksheet
     *
     * @param name the worksheet name
     */
    public ListMapSheet(String name) {
        super(name);
    }

    /**
     * Constructor worksheet
     *
     * @param columns the header info
     */
    public ListMapSheet(final Column... columns) {
        super(columns);
    }

    /**
     * Constructor worksheet
     *
     * @param name    the worksheet name
     * @param columns the header info
     */
    public ListMapSheet(String name, final Column... columns) {
        super(name, columns);
    }

    /**
     * Constructor worksheet
     *
     * @param name      the worksheet name
     * @param waterMark the water mark
     * @param columns   the header info
     */
    public ListMapSheet(String name, WaterMark waterMark, final Column... columns) {
        super(name, waterMark, columns);
    }


    /**
     * Constructor worksheet
     *
     * @param data the worksheet's body data
     */
    public ListMapSheet(List<Map<String, ?>> data) {
        this(null, data);
    }

    /**
     * Constructor worksheet
     *
     * @param name the worksheet name
     * @param data the worksheet's body data
     */
    public ListMapSheet(String name, List<Map<String, ?>> data) {
        super(name);
        setData(data);
    }

    /**
     * Constructor worksheet
     *
     * @param data    the worksheet's body data
     * @param columns the header info
     */
    public ListMapSheet(List<Map<String, ?>> data, final Column... columns) {
        this(null, data, columns);
    }

    /**
     * Constructor worksheet
     *
     * @param name    the worksheet name
     * @param data    the worksheet's body data
     * @param columns the header info
     */
    public ListMapSheet(String name, List<Map<String, ?>> data, final Column... columns) {
        this(name, data, null, columns);
    }

    /**
     * Constructor worksheet
     *
     * @param data      the worksheet's body data
     * @param waterMark the water mark
     * @param columns   the header info
     */
    public ListMapSheet(List<Map<String, ?>> data, WaterMark waterMark, final Column... columns) {
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
    public ListMapSheet(String name, List<Map<String, ?>> data, WaterMark waterMark, final Column... columns) {
        super(name, waterMark, columns);
        setData(data);
    }

    /**
     * Reset the row-block data
     */
    @Override
    protected void resetBlockData() {
        if (!eof && left() < rowBlock.capacity()) {
            append();
        }
        int end = getEndIndex(), len = columns.length;
        boolean hasGlobalStyleProcessor = (extPropMark & 2) == 2;
        for (; start < end; rows++, start++) {
            Row row = rowBlock.next();
            row.index = rows;
            row.height = getRowHeight();
            Cell[] cells = row.realloc(len);
            Map<String, ?> rowDate = data.get(start);
            for (int i = 0; i < len; i++) {
                Column hc = columns[i];
                Object e = rowDate != null ? rowDate.get(hc.key) : null;
                // Clear cells
                Cell cell = cells[i];
                cell.clear();

                cellValueAndStyle.reset(row, cell, e, hc);
                if (hasGlobalStyleProcessor) {
                    cellValueAndStyle.setStyleDesign(rowDate, cell, hc, getStyleProcessor());
                }
            }
        }
    }


    /**
     * Returns the header column info
     *
     * @return array of column
     */
    @Override
    protected Column[] getHeaderColumns() {
        if (headerReady) return columns;
        Map<String, ?> first = getFirst();
        // No data
        if (first == null) {
            if (columns == null) {
                columns = new Column[0];
            }
        } else if (!hasHeaderColumns()) {
            int size = first.size(), i = 0;
            columns = new Column[size];
            for (Iterator<? extends Map.Entry<String, ?>> it = first.entrySet().iterator(); it.hasNext(); ) {
                Map.Entry<String, ?> entry = it.next();
                Column hc = createColumn(entry);
                if (hc != null) columns[i++] = hc;
            }
            if (i < size) columns = Arrays.copyOf(columns, i);
        } else {
            Object o;
            for (int i = 0; i < columns.length; i++) {
                Column hc = columns[i].getTail();
                if (isEmpty(hc.key)) {
                    throw new ExcelWriteException(getClass() + " must specify the 'key' name.");
                }
                if (hc.getClazz() == null) {
                    hc.setClazz((o = first.get(hc.key)) != null ? o.getClass() : String.class);
                }
            }
        }

        return columns;
    }

    /**
     * Create column from {@link Map.Entry}
     * <p>
     * Override the method to extend custom comments
     *
     * @param entry the first entry from ListMap
     * @return the Worksheet's {@link EntryColumn} information
     */
    protected Column createColumn(Map.Entry<String, ?> entry) {
        // Ignore the null key
        if (isEmpty(entry.getKey())) return null;
        Object value = entry.getValue();
        return new Column(entry.getKey(), entry.getKey(), value != null ? value.getClass() : String.class);
    }

    /**
     * Merge cells if
     */
    @Override
    protected void mergeHeaderCellsIfEquals() {
        super.mergeHeaderCellsIfEquals();

        @SuppressWarnings("unchecked")
        List<Dimension> existsMergeCells = (List<Dimension>) getExtPropValue(Const.ExtendPropertyKey.MERGE_CELLS);
        if (existsMergeCells != null) {
            Grid grid = GridFactory.create(existsMergeCells);
            for (Column col : columns) {
                if (StringUtil.isEmpty(col.key) && grid.test(1, col.realColIndex)) {
                    Column next = col.next;
                    for (; next != null && StringUtil.isEmpty(next.key); next = next.next);
                    if (next != null) col.key = next.key; // Keep the key to get the value
                }
            }
        }
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

        // As default as force export
        resetBlockData();

        return rowBlock.flip();
    }
}

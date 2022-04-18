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

import org.ttzero.excel.reader.Cell;

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
     * @param name    the worksheet name
     * @param columns the header info
     */
    public ListMapSheet(String name, final org.ttzero.excel.entity.Column... columns) {
        super(name, columns);
    }

    /**
     * Constructor worksheet
     *
     * @param name      the worksheet name
     * @param waterMark the water mark
     * @param columns   the header info
     */
    public ListMapSheet(String name, WaterMark waterMark, final org.ttzero.excel.entity.Column... columns) {
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
    public ListMapSheet(List<Map<String, ?>> data, final org.ttzero.excel.entity.Column... columns) {
        this(null, data, columns);
    }

    /**
     * Constructor worksheet
     *
     * @param name    the worksheet name
     * @param data    the worksheet's body data
     * @param columns the header info
     */
    public ListMapSheet(String name, List<Map<String, ?>> data, final org.ttzero.excel.entity.Column... columns) {
        this(name, data, null, columns);
    }

    /**
     * Constructor worksheet
     *
     * @param data      the worksheet's body data
     * @param waterMark the water mark
     * @param columns   the header info
     */
    public ListMapSheet(List<Map<String, ?>> data, WaterMark waterMark, final org.ttzero.excel.entity.Column... columns) {
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
    public ListMapSheet(String name, List<Map<String, ?>> data, WaterMark waterMark, final org.ttzero.excel.entity.Column... columns) {
        super(name, waterMark, columns);
        setData(data);
    }

    /**
     * Reset the row-block data
     */
    @Override
    protected void resetBlockData() {
        if (!eof && left() < getRowBlockSize()) {
            append();
        }
        int end = getEndIndex();
        int len = columns.length;
        for (; start < end; rows++, start++) {
            Row row = rowBlock.next();
            row.index = rows;
            Cell[] cells = row.realloc(len);
            for (int i = 0; i < len; i++) {
                org.ttzero.excel.entity.Column hc = columns[i];
                Object e = data.get(start).get(hc.key);
                // clear cells
                Cell cell = cells[i];
                cell.clear();

                cellValueAndStyle.reset(rows, cell, e, hc);
                cellValueAndStyle.setStyleDesign(data.get(start), cell, columns[i], getStyleProcessor());
            }
        }
    }


    /**
     * Returns the header column info
     *
     * @return array of column
     */
    @Override
    protected org.ttzero.excel.entity.Column[] getHeaderColumns() {
        if (headerReady) return columns;
        Map<String, ?> first = getFirst();
        // No data
        if (first == null) {
            if (columns == null) {
                columns = new org.ttzero.excel.entity.Column[0];
            }
        } else if (!hasHeaderColumns()) {
            int size = first.size(), i = 0;
            columns = new org.ttzero.excel.entity.Column[size];
            for (Iterator<? extends Map.Entry<String, ?>> it = first.entrySet().iterator(); it.hasNext(); ) {
                Map.Entry<String, ?> entry = it.next();
                // Ignore the null key
                if (isEmpty(entry.getKey())) continue;
                Object value = entry.getValue();
                columns[i++] = new org.ttzero.excel.entity.Column(entry.getKey(), entry.getKey(), value != null ? value.getClass() : String.class);
            }
        } else {
            for (int i = 0; i < columns.length; i++) {
                org.ttzero.excel.entity.Column hc = columns[i];
                if (isEmpty(hc.key)) {
                    throw new ExcelWriteException(getClass() + " must specify the 'key' name.");
                }
                if (hc.getClazz() == null) {
                    hc.setClazz(first.get(hc.key).getClass());
                }
            }
        }

        return columns;
    }

}

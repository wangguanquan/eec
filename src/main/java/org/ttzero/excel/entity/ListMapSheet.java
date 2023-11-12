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
 * {@code ListMapSheet}是{@code ListSheet}的一个子集，因为取值方式完全不同所以将其独立，
 * 未指定表头信息时{@code ListMapSheet}将导出全字段这是与{@code ListSheet}完全不同的设定
 *
 * @author guanquan.wang at 2018-01-26 14:46
 * @see ListSheet
 */
public class ListMapSheet extends ListSheet<Map<String, ?>> {

    /**
     * 实例化工作表，未指定工作表名称时默认以{@code 'Sheet'+id}命名
     */
    public ListMapSheet() {
        super();
    }

    /**
     * 实例化工作表并指定工作表名称
     *
     * @param name 工作表名称
     */
    public ListMapSheet(String name) {
        super(name);
    }

    /**
     * 实例化工作表并指定表头信息
     *
     * @param columns 表头信息
     */
    public ListMapSheet(final Column... columns) {
        super(columns);
    }

    /**
     * 实例化工作表并指定工作表名称和表头信息
     *
     * @param name    工作表名称
     * @param columns 表头信息
     */
    public ListMapSheet(String name, final Column... columns) {
        super(name, columns);
    }

    /**
     * 实例化工作表并指定工作表名称，水印和表头信息
     *
     * @param name      工作表名称
     * @param waterMark 水印
     * @param columns   表头信息
     */
    public ListMapSheet(String name, WaterMark waterMark, final Column... columns) {
        super(name, waterMark, columns);
    }

    /**
     * 实例化工作表并添加导出数据
     *
     * @param data 需要导出的数据
     */
    public ListMapSheet(List<Map<String, ?>> data) {
        this(null, data);
    }

    /**
     * 实例化工作表并指定工作表名和添加导出数据
     *
     * @param name 工作表名称
     * @param data 需要导出的数据
     */
    public ListMapSheet(String name, List<Map<String, ?>> data) {
        super(name);
        setData(data);
    }

    /**
     * 实例化工作表并添加导出数据和表头信息
     *
     * @param data    需要导出的数据
     * @param columns 表头信息
     */
    public ListMapSheet(List<Map<String, ?>> data, final Column... columns) {
        this(null, data, columns);
    }

    /**
     * 实例化指定名称工作表并添加导出数据和表头信息
     *
     * @param name    工作表名称
     * @param data    需要导出的数据
     * @param columns 表头信息
     */
    public ListMapSheet(String name, List<Map<String, ?>> data, final Column... columns) {
        this(name, data, null, columns);
    }

    /**
     * 实例化工作表并添加导出数据和表头信息
     *
     * @param data      需要导出的数据
     * @param waterMark 水印
     * @param columns   表头信息
     */
    public ListMapSheet(List<Map<String, ?>> data, WaterMark waterMark, final Column... columns) {
        this(null, data, waterMark, columns);
    }

    /**
     * 实例化指定名称工作表并添加导出数据和表头信息
     *
     * @param name      工作表名
     * @param data      需要导出的数据
     * @param waterMark 水印
     * @param columns   表头信息
     */
    public ListMapSheet(String name, List<Map<String, ?>> data, WaterMark waterMark, final Column... columns) {
        super(name, waterMark, columns);
        setData(data);
    }

    /**
     * 重置{@code RowBlock}行块数据
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
     * 获取表头信息，未指定{@code Columns}时默认以第1个非{@code null}的Map值做为参考，将该map中所有key做为表头
     *
     * @return 初如化表头
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
                boolean emptyKey = isEmpty(hc.key), emptyName = isEmpty(hc.name);
                if (emptyKey && emptyName) throw new ExcelWriteException(getClass() + " must specify the 'key' name.");
                else if (emptyKey) hc.key = hc.name;
                else if (emptyName) hc.name = hc.key;
                if (hc.getClazz() == null) {
                    hc.setClazz((o = first.get(hc.key)) != null ? o.getClass() : String.class);
                }
            }
        }

        return columns;
    }

    /**
     * 从{@link Map.Entry}提取信息创建表头，除忽略{@code null}作为key以外的其它所有key均默认导出
     *
     * @param entry 第一个非{@code null}的map包含的所有值
     * @return 单列表头信息
     */
    protected Column createColumn(Map.Entry<String, ?> entry) {
        // Ignore the null key
        if (isEmpty(entry.getKey())) return null;
        Object value = entry.getValue();
        return new Column(entry.getKey(), entry.getKey(), value != null ? value.getClass() : String.class);
    }

    /**
     * 合并表头
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
                    for (; next != null && StringUtil.isEmpty(next.key); next = next.next) ;
                    if (next != null) col.key = next.key; // Keep the key to get the value
                }
            }
        }
    }

    /**
     * 获取下一段{@link RowBlock}行块数据，工作表输出协议通过此方法循环获取行数据并落盘，
     * 行块被设计为一个滑行窗口，下游输出协议只能获取一个窗口的数据默认包含32行。
     *
     * @return 行块
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

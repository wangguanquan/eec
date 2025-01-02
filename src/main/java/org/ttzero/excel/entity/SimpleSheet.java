/*
 * Copyright (c) 2017-2024, guanquan.wang@hotmail.com All Rights Reserved.
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

import org.ttzero.excel.entity.style.Horizontals;
import org.ttzero.excel.entity.style.NumFmt;
import org.ttzero.excel.entity.style.Styles;
import org.ttzero.excel.reader.Cell;

import java.lang.reflect.Array;
import java.util.Arrays;
import java.util.List;

/**
 * 简单工作表，它的简单之处在于只需要指定单元格的值即可输出，不需要再定义对象也不受&#x40;ExcelColumn注解影响
 * 例如{@code Arrays.asList(1, 2, 3)}则会将数字{@code 1,2,3}输出到第一行的{@code A,B,C}三列。
 *
 * <p>简单工作表支持传入{@code List}或{@code Array}，但两种类型不能掺杂使用，要么全是{@code List}
 * 要么全是{@code Array}，导出过程中并不会逐一判断泛型的实际类型，所以需要在外部做好约束。由于没有表头做为限制所以每一行的数组长度
 * 并不要求一致数据类型也不要求一致但最长不能超过Excel限制的长度，除日期类型外每个单元格均保持"常规"格式，日期类型默认添加格式化。
 * {@code SimpleSheet}继承于{@code ListSheet}如果传入数据为普通数组对象则将使用{@code ListSheet}进行处理。</p>
 *
 * <p>默认情况下简单工作表将不会设置样式(包括表头样式)，{@code SimpleSheet}提供简化的{@link #setHeader(List)}方法来指定表头，
 * 也可以使用{@link #firstRowAsHeader}方法将第一行数据做为表头，当指定表头时依然会保持通用样式。</p>
 *
 * <pre>
 * new Workbook()
 *     .addSheet(new SimpleSheet&lt;&gt;()
 *          // 导出两行数据
 *         .setData(Arrays.asList(new String[]{"a","b","c"}, new int[]{1,2,3,4,5}))
 *     ).writeTo(Paths.get("f://abc.xlsx"));</pre>
 * @author guanquan.wang
 * @since 2024-09-24
 * @see ListSheet
 */
public class SimpleSheet<T> extends ListSheet<T> {
    /**
     * 0: empty 1: List 2: Array 3: Super(not a type)
     */
    protected int type = -1;
    /**
     * 将第一行数据作为表头
     */
    protected boolean firstRowAsHeader;
    /**
     * 未实例化的列，可用于在写超出预知范围外的列
     */
    protected static final Column UNALLOCATED_COLUMN = new Column();
    /**
     * 实例化工作表，未指定工作表名称时默认以{@code 'Sheet'+id}命名
     */
    public SimpleSheet() { }
    /**
     * 实例化工作表并指定工作表名称
     *
     * @param name 工作表名称
     */
    public SimpleSheet(String name) {
        super(name);
    }
    /**
     * 实例化工作表并指定表头信息
     *
     * @param columns 表头信息
     */
    public SimpleSheet(Column... columns) {
        super(columns);
    }
    /**
     * 实例化工作表并指定工作表名称和表头信息
     *
     * @param name    工作表名称
     * @param columns 表头信息
     */
    public SimpleSheet(String name, Column... columns) {
        super(name, columns);
    }
    /**
     * 实例化工作表并指定工作表名称，水印和表头信息
     *
     * @param name      工作表名称
     * @param waterMark 水印
     * @param columns   表头信息
     */
    public SimpleSheet(String name, WaterMark waterMark, Column... columns) {
        super(name, waterMark, columns);
    }
    /**
     * 实例化工作表并指定初始数据
     *
     * @param data 初始数据
     */
    public SimpleSheet(List<T> data) {
        super(data);
    }
    /**
     * 实例化工作表并指定工作表名称和初始数据
     *
     * @param name 工作表名称
     * @param data 初始数据
     */
    public SimpleSheet(String name, List<T> data) {
        super(name, data);
    }
    /**
     * 实例化工作表并指定初始数据和表头
     *
     * @param data    初始数据
     * @param columns 表头信息
     */
    public SimpleSheet(List<T> data, Column... columns) {
        super(data, columns);
    }
    /**
     * 实例化工作表并指定工作表名称、初始数据和表头
     *
     * @param name    工作表名称
     * @param data    初始数据
     * @param columns 表头信息
     */
    public SimpleSheet(String name, List<T> data, Column... columns) {
        super(name, data, columns);
    }
    /**
     * 实例化工作表并指定初始数据、水印和表头
     *
     * @param data      初始数据
     * @param waterMark 水印
     * @param columns   表头信息
     */
    public SimpleSheet(List<T> data, WaterMark waterMark, Column... columns) {
        super(data, waterMark, columns);
    }
    /**
     * 实例化工作表并指定工作表名称、初始数据、水印和表头
     *
     * @param name      工作表名称
     * @param data      初始数据
     * @param waterMark 水印
     * @param columns   表头信息
     */
    public SimpleSheet(String name, List<T> data, WaterMark waterMark, Column... columns) {
        super(name, data, waterMark, columns);
    }

    /**
     * 设置表头信息，与Columns不同的是本方法只设置表头值并不带任何其它属性，可以看为{@link #setColumns(List)}的简化方法
     *
     * @param header 表头信息列表
     * @return 当前对象，支持链式调用
     */
    public SimpleSheet<T> setHeader(List<String> header) {
        Column[] columns;
        if (header == null || header.isEmpty()) columns = new Column[0];
        else {
            columns = new Column[header.size()];
            for (int i = 0, len = header.size(); i < len; columns[i] = new Column(header.get(i++)).setCellStyle(0));
        }
        super.setColumns(columns);
        return this;
    }

    /**
     * 设置表头信息，与Columns不同的是本方法只设置表头值并不带任何其它属性，可以看为{@link #setColumns(Column...)}的简化方法
     *
     * @param header 表头信息列表
     * @return 当前对象，支持链式调用
     */
    public SimpleSheet<T> setHeader(String ... header) {
        return setHeader(Arrays.asList(header));
    }

    /**
     * 将第一行数据作为表头
     *
     * @return SimpleSheet对象，包含了表头信息
     */
    public SimpleSheet<T> firstRowAsHeader() {
        firstRowAsHeader = true;
        return this;
    }

    /**
     * 获取表头信息，未实例化表头时会执行初始化方法实例化表头
     *
     * <p>对于简单类型来说表头信息并无任何有效信息，</p>
     *
     * @return 表头信息
     */
    @Override
    public Column[] getHeaderColumns() {
        Object o = getFirst();
        if (o == null) type = 0;
        // List
        else if (List.class.isAssignableFrom(o.getClass())) {
            type = 1;
            // 将第一行做为头表
            if (firstRowAsHeader) {
                List row0 = (List) o;
                columns = new Column[row0.size()];
                int i = 0;
                for (Object e : row0) columns[i++] = new Column(e.toString()).setCellStyle(0);
                // 这里取了第一行所以将start向前移动
                start++;
            }
        }
        // 数组
        else if (o.getClass().isArray()) {
            type = 2;
            // 将第一行做为头表
            if (firstRowAsHeader) {
                int len = Array.getLength(o);
                columns = new Column[len];
                for (int i = 0; i < len; i++) columns[i] = new Column(Array.get(o, i).toString()).setCellStyle(0);
                // 这里取了第一行所以将start向前移动
                start++;
            }
        // 普通ListSheet
        } else {
            type = 3;
            return super.getHeaderColumns();
        }

        // 特殊设置
        if (columns == null) {
            columns = new Column[0];
            headerReady = true;
            // 默认忽略表头
            ignoreHeader();
            setHeaderRowHeight(-1D);
        }
        UNALLOCATED_COLUMN.styles = workbook.getStyles();
        UNALLOCATED_COLUMN.cellStyle = 0; // General Style

        return columns;
    }

    /**
     * 重置{@code RowBlock}行块数据
     */
    @Override
    protected void resetBlockData() {
        // 普通ListSheet
        if (type == 3) {
            super.resetBlockData();
            return;
        }

        if (!eof && left() < rowBlock.capacity()) append();

        // Find the end index of row-block
        int end = getEndIndex();
        for (; start < end; rows++, start++) {
            Row row = rowBlock.next();
            row.index = rows;
            T o = data.get(start);
            boolean isNull = o == null;
            List sub = !isNull && type == 1 ? (List) o : null;
            int len = !isNull ? type == 1 ? sub.size() : Array.getLength(o) : 0;
            Cell[] cells = row.realloc(len);
            for (int i = 0; i < len; i++) {
                // Clear cells
                Cell cell = cells[i];
                cell.clear();

                Object e = null;
                Column column = i < columns.length ? columns[i] : UNALLOCATED_COLUMN;
                // 根据下标取数
                if (!column.isIgnoreValue()) {
                    if (type == 1) e = sub.get(i);
                    else e = Array.get(o, i);
                }
                column.clazz = null; // 无法确定纵向类型完全一致所以这里将缓存的类型清除
                cellValueAndStyle.reset(row, cell, e, column);
                // 日期类型添加默认format
                if (cell.t == Cell.DATETIME || cell.t == Cell.DATE || cell.t == Cell.TIME) {
                    datetimeCell(workbook.getStyles(), cell);
                }
            }
            row.height = getRowHeight();
        }
    }

    /**
     * 日期类型添加默认format
     *
     * @param styles Styles
     * @param cell 单元格
     */
    protected static void datetimeCell(Styles styles, Cell cell) {
        defaultDatetimeCell(styles, cell);
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

    /**
     * 获取默认列宽，如果未在Column上特殊指定宽度时该宽度将应用于每一列
     *
     * @return 默认列宽10
     */
    @Override
    public double getDefaultWidth() {
        return type <= 2 ? 10.16D : super.getDefaultWidth();
    }

    /**
     * 日期类型添加默认format
     *
     * @param styles Styles
     * @param cell 单元格
     */
    public static void defaultDatetimeCell(Styles styles, Cell cell) {
        // 已有日期格式化则保留
        if (styles.isDate(cell.xf)) return;
        int style = styles.getStyleByIndex(cell.xf);
        switch (cell.t) {
            case Cell.DATETIME: style = styles.modifyNumFmt(style, NumFmt.DATETIME_FORMAT); break;
            case Cell.DATE    : style = styles.modifyNumFmt(style, NumFmt.DATE_FORMAT);     break;
            case Cell.TIME    : style = styles.modifyNumFmt(style, NumFmt.TIME_FORMAT);     break;
        }
        if (!Styles.hasHorizontal(style)) style |= Horizontals.CENTER;
        cell.xf = styles.of(style);
    }
}

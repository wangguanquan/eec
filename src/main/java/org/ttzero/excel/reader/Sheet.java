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

import org.ttzero.excel.manager.Const;
import org.ttzero.excel.util.CSVUtil;
import org.ttzero.excel.util.FileUtil;
import org.ttzero.excel.util.StringUtil;

import java.io.BufferedWriter;
import java.io.Closeable;
import java.io.IOException;
import java.io.OutputStream;
import java.io.OutputStreamWriter;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Iterator;
import java.util.List;
import java.util.Spliterator;
import java.util.Spliterators;
import java.util.stream.Stream;
import java.util.stream.StreamSupport;

import static org.ttzero.excel.reader.Cell.BOOL;
import static org.ttzero.excel.reader.Cell.DECIMAL;
import static org.ttzero.excel.reader.Cell.DOUBLE;
import static org.ttzero.excel.reader.Cell.FUNCTION;
import static org.ttzero.excel.reader.Cell.INLINESTR;
import static org.ttzero.excel.reader.Cell.LONG;
import static org.ttzero.excel.reader.Cell.NUMERIC;
import static org.ttzero.excel.reader.Cell.SST;
import static org.ttzero.excel.util.DateUtil.toLocalDate;
import static org.ttzero.excel.util.DateUtil.toTimestamp;
import static org.ttzero.excel.util.FileUtil.exists;

/**
 * 用于读的工作表，为了性能本工具将读和写分开设计它们具有完全不同的方法，
 * 读取数据时可以通过{@link #header}方法指定表头位置
 *
 * @author guanquan.wang at 2019-04-17 11:36
 */
public interface Sheet extends Closeable {

    /**
     * 获取工作表的名称
     *
     * @return 工作表名
     */
    String getName();

    /**
     * The index of worksheet located at the workbook
     *
     * @return the index(zero base)
     */
    int getIndex();

    /**
     * The worksheet id, it difference with index is that the id will not change
     * because of moving or deleting worksheet.
     *
     * @return id of worksheet
     */
    int getId();

    /**
     * size of rows.
     *
     * @return size of rows
     *      -1: unknown size
     * @deprecated use {@link #getDimension()} to getting full range address
     */
    @Deprecated
    int getSize();

    /**
     * Returns The range address of the used area in
     * the current sheet
     * <p>
     * NOTE: This method can only guarantee accurate row ranges
     *
     * @return worksheet {@link Dimension} ranges
     */
    Dimension getDimension();

    /**
     * Test Worksheet is hidden
     *
     * @return true if current worksheet is hidden
     */
    boolean isHidden();

    /**
     * Test Worksheet is show
     *
     * @return true if current worksheet is show
     */
    default boolean isShow() {
        return !isHidden();
    }

    /**
     * Specify the header rows endpoint
     *
     * @param fromRowNum low endpoint (inclusive) of the worksheet (one base)
     * @return current {@link Sheet}
     * @throws IndexOutOfBoundsException if {@code fromRow} less than 1
     */
    default Sheet header(int fromRowNum) {
        return header(fromRowNum, fromRowNum);
    }

    /**
     * Specify the header rows endpoint
     * <p>
     * Note: After specifying the header row number, the row-pointer will move to the
     * next row of the header range. The {@link #bind(Class)}, {@link #bind(Class, int)},
     * {@link #bind(Class, int, int)}, {@link #rows()}, {@link #dataRows()}, {@link #iterator()},
     * and {@link #dataIterator()} will all be affected.
     *
     * @param fromRowNum low endpoint (inclusive) of the worksheet (one base)
     * @param toRowNum high endpoint (inclusive) of the worksheet (one base)
     * @return current {@link Sheet}
     * @throws IndexOutOfBoundsException if {@code fromRow} less than 1
     * @throws IllegalArgumentException if {@code toRow} less than {@code fromRow}
     */
    Sheet header(int fromRowNum, int toRowNum);

    /**
     * Returns the header of the list.
     *
     * The first non-empty line defaults to the header information. You can also call {@link #header(int, int)}
     * to specify multiple header rows. If there are multiple rows of headers, ':' will be used for stitching.
     *
     * <blockquote><pre>
     * +-----------------------------+
     * |       |        COMMON       |
     * | TITLE +-------+------+------+
     * |       |  SUB1 | SUB2 | SUB3 |
     * +------+-------+-------+------+
     * </pre></blockquote>
     *
     * The above table will return "TITLE", "COMMON:SUB1", "COMMON:SUB2", "COMMON:SUB3"
     *
     * @return the {@link HeaderRow}
     */
    Row getHeader();

    /**
     * Set the binding type
     *
     * @param clazz the binding type
     * @return the {@link Sheet}
     */
    Sheet bind(Class<?> clazz);

    /**
     * Set the binding type
     *
     * @param clazz the binding type
     * @param fromRowNum low endpoint (inclusive) of the worksheet (one base)
     * @return the {@link Sheet}
     */
    default Sheet bind(Class<?> clazz, int fromRowNum) {
        return bind(clazz, header(fromRowNum).getHeader());
    }

    /**
     * Set the binding type
     *
     * @param clazz the binding type
     * @param fromRowNum low endpoint (inclusive) of the worksheet (one base)
     * @param toRowNum high endpoint (inclusive) of the worksheet (one base)
     * @return the {@link Sheet}
     */
    default Sheet bind(Class<?> clazz, int fromRowNum, int toRowNum) {
        return bind(clazz, header(fromRowNum, toRowNum).getHeader());
    }

    /**
     * Set the binding type
     *
     * @param clazz the binding type
     * @param row specify a custom header row
     * @return the {@link Sheet}
     */
    Sheet bind(Class<?> clazz, Row row);

    /**
     * Load the sheet data
     *
     * @return the {@link Sheet}
     * @throws IOException if I/O error occur
     */
    Sheet load() throws IOException;

    /**
     * Iterating each row of data contains header information and blank lines
     *
     * @return a row iterator
     */
    Iterator<Row> iterator();

    /**
     * Iterating over data rows without header information and blank lines
     *
     * @return a row iterator
     */
    Iterator<Row> dataIterator();

    /**
     * List all pictures in workbook
     *
     * @return picture list or null if not exists.
     */
    List<Drawings.Picture> listPictures();

    /**
     * Reset the {@link Sheet}'s row index to begging
     *
     * @return the unread {@link Sheet}
     * @throws ExcelReadException if I/O error occur.
     * @throws UnsupportedOperationException if sub-class un-implement this function.
     */
    default Sheet reset() {
        throw new UnsupportedOperationException();
    }

    /**
     * Return a stream of all rows
     *
     * @return a {@code Stream&lt;Row&gt;} providing the lines of row
     * described by this {@link Sheet}
     * @since 1.8
     */
    default Stream<Row> rows() {
        return StreamSupport.stream(Spliterators.spliteratorUnknownSize(
            iterator(), Spliterator.ORDERED | Spliterator.NONNULL), false);
    }

    /**
     * Return stream with out header row and empty rows
     *
     * @return a {@code Stream&lt;Row&gt;} providing the lines of row
     * described by this {@link Sheet}
     * @since 1.8
     */
    default Stream<Row> dataRows() {
        return StreamSupport.stream(Spliterators.spliteratorUnknownSize(
            dataIterator(), Spliterator.ORDERED | Spliterator.NONNULL), false);
    }

    /**
     * Convert column mark to int
     *
     * @param col column mark
     * @return int value
     */
    static int col2Int(String col) {
        if (StringUtil.isEmpty(col)) return 1;
        char[] values = col.toCharArray();
        int n = 0;
        for (char value : values) {
            if (value < 'A' || value > 'Z')
                throw new ExcelReadException("Column mark out of range: " + col);
            n = n * 26 + value - 'A' + 1;
        }
        return n;
    }

//    /**
//     * Close resource
//     *
//     * @throws IOException if I/O error occur
//     */
//    @Override
//    void close() throws IOException;

    /**
     * Save file as Comma-Separated Values. Each worksheet corresponds to
     * a csv file. Default charset is 'UTF8' and separator character is ','.
     *
     * @param path the output storage path
     * @throws IOException if I/O error occur.
     */
    default void saveAsCSV(Path path) throws IOException {
        saveAsCSV(path, StandardCharsets.UTF_8);
    }

    /**
     * Save file as Comma-Separated Values. Each worksheet corresponds to
     * a csv file.
     *
     * @param path the output storage path
     * @param charset specify a charset, default is UTF-8
     * @throws IOException if I/O error occur.
     */
    default void saveAsCSV(Path path, Charset charset) throws IOException {
        // Create path if not exists
        if (!exists(path)) {
            FileUtil.mkdir(path);
        }
        if (Files.isDirectory(path)) {
            path = path.resolve(getName() + Const.Suffix.CSV);
        }

        saveAsCSV(Files.newOutputStream(path), charset);
    }

    /**
     * Save file as Comma-Separated Values. Each worksheet corresponds to
     * a csv file. Default charset is 'UTF8' and separator character is ','.
     *
     * @param os the output
     * @throws IOException if I/O error occur.
     */
    default void saveAsCSV(OutputStream os) throws IOException {
        saveAsCSV(os, StandardCharsets.UTF_8);
    }

    /**
     * Save file as Comma-Separated Values. Each worksheet corresponds to
     * a csv file. Default separator character is ','.
     *
     * @param os the output
     * @param charset specify a charset
     * @throws IOException if I/O error occur.
     */
    default void saveAsCSV(OutputStream os, Charset charset) throws IOException {
        saveAsCSV(new BufferedWriter(new OutputStreamWriter(os, charset)));
    }

    /**
     * Save file as Comma-Separated Values. Each worksheet corresponds to
     * a csv file. Default charset is 'UTF8' and separator character is ','.
     *
     * @param bw buffer writer
     * @throws IOException if I/O error occur.
     */
    default void saveAsCSV(BufferedWriter bw) throws IOException {
        try (CSVUtil.Writer writer = CSVUtil.newWriter(bw)) {
            for (Iterator<Row> iter = iterator(); iter.hasNext(); ) {
                Row row = iter.next();
                if (row.isEmpty()) {
                    writer.newLine();
                    continue;
                }
                for (int i = 0; i < row.lc; i++) {
                    Cell c = row.cells[i];
                    switch (c.t) {
                        case SST       : if (c.sv == null) c.setSv(row.sst.get(c.nv));
                        case INLINESTR :
                        case FUNCTION  : writer.write(c.sv); break;
                        case NUMERIC   :
                            if (!row.styles.fastTestDateFmt(c.xf)) writer.write(c.nv);
                            else writer.write(toLocalDate(c.nv).toString());
                            break;
                        case LONG      : writer.write(c.lv); break;
                        case DECIMAL   :
                            if (!row.styles.fastTestDateFmt(c.xf)) writer.write(c.mv.toString());
                            else writer.write(toTimestamp(c.mv.doubleValue()).toString());
                            break;
                        case DOUBLE    :
                            if (!row.styles.fastTestDateFmt(c.xf)) writer.write(c.dv);
                            else writer.write(toTimestamp(c.dv).toString());
                            break;
                        case BOOL      : writer.write(c.bv); break;
                        default        : writer.writeEmpty();
                    }
                }
                writer.newLine();
            }
        }
    }

    /**
     * Use field name matching without {@link org.ttzero.excel.annotation.ExcelColumn} annotation
     *
     * <p>When converting row data to Java objects, only fields with {@code ExcelColumn} annotations
     * are matched by default. {@code forceImport} will skip this restriction,
     * and fields without {@code ExcelColumn} annotations will be matched with field name
     *
     * @return 当前工作表
     */
    default Sheet forceImport() {
        return addHeaderColumnReadOption(HeaderRow.FORCE_IMPORT);
    }

    /**
     * 设置忽略大小写匹配表头字段
     *
     * @return 当前工作表
     */
    default Sheet headerColumnIgnoreCase() {
        return addHeaderColumnReadOption(HeaderRow.IGNORE_CASE);
    }

    /**
     * 设置驼峰风格匹配表头字段
     *
     * @return 当前工作表
     */
    default Sheet headerColumnToCamelCase() {
        return addHeaderColumnReadOption(HeaderRow.CAMEL_CASE);
    }

    /**
     * 添加表头读取属性
     *
     * @param option 额外属性
     * @return 当前工作表
     */
    default Sheet addHeaderColumnReadOption(int option) {
        return setHeaderColumnReadOption(getHeaderColumnReadOption() | option);
    }

    /**
     * 设置表头读取属性，将行数据转对象时由于Excel中的值与Java对象中定义的不同会使双方不匹配，设置读取属性可丰富读取能力，
     * 多个属性可叠加
     *
     * <ul>
     *     <li>HeaderRow.FORCE_IMPORT: 强制导入，即使没有&#40;ExcelColumn注解</li>
     *     <li>HeaderRow.IGNORE_CASE: 忽略大小写匹配</li>
     *     <li>HeaderRow.CAMEL_CASE: 驼峰风格匹配</li>
     * </ul>
     * <blockquote><pre>
     *     reader.sheet(0).setHeaderColumnReadOption(HeaderRow.FORCE_IMPORT | HeaderRow.IGNORE_CASE)
     * </pre></blockquote>
     *
     * @param option 额外属性
     * @return 当前工作表
     */
    Sheet setHeaderColumnReadOption(int option);

    /**
     * 获取表头读取属性
     *
     * @return 属性值
     */
    int getHeaderColumnReadOption();

    /**
     * 将工作表转为普通工作表{@code Sheet}，它只专注获取值
     *
     * @return {@link Sheet}
     */
    Sheet asSheet();

    /**
     * 将工作表转为{@code CalcSheet}以解析单元格公式
     *
     * @return {@link CalcSheet}
     */
    CalcSheet asCalcSheet();

    /**
     * 将工作表转为{@code MergeSheet}，它将复制合并单元格的首坐标值到合并范围内的其它单元格中
     *
     * @return {@link MergeSheet}
     */
    MergeSheet asMergeSheet();

    /**
     * 将工作表转为{@code FullSheet}支持全属性读取
     *
     * @return {@link FullSheet}
     */
    FullSheet asFullSheet();
}
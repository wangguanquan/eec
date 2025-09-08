/*
 * Copyright (c) 2017-2019, guanquan.wang@hotmail.com All Rights Reserved.
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

import java.io.BufferedWriter;
import java.io.Closeable;
import java.io.IOException;
import java.io.OutputStream;
import java.io.OutputStreamWriter;
import java.math.BigDecimal;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalTime;
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
import static org.ttzero.excel.util.DateUtil.toTimeChars;
import static org.ttzero.excel.util.DateUtil.toDateTimeString;
import static org.ttzero.excel.util.DateUtil.toLocalDate;
import static org.ttzero.excel.util.DateUtil.toLocalTime;
import static org.ttzero.excel.util.DateUtil.toTimestamp;

/**
 * 用于读的工作表，为了性能本工具将读和写分开设计它们具有完全不同的方法，
 * 读取数据时可以通过{@link #header}方法指定表头位置，多行表头时可以指定一个起始行和结束行
 * 来匹配，它将以{@code 行1:行2...行n}拼按的形式做为Key
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
     * 获取工作表在工作薄中的下标（从0开始）
     *
     * @return 工作表下标(从0开始)
     */
    int getIndex();

    /**
     * 工作表id，它与索引的区别在于，id不会因为移动或删除工作表而更改。
     *
     * @return 工作表id
     */
    int getId();

    /**
     * 获取当前工作表的总行数
     *
     * @return 当前工作表的总行数，{@code -1}表示未知
     */
    default int getSize() {
        Dimension d = getDimension();
        return d != null ? d.lastRow - d.firstRow + 1 : -1;
    }

    /**
     * 获取当前工作表中有效区域的范围地址（任意值，样式均表示有效值），此值取于头信息&lt;dimension&gt;的值，
     * 如果头信息没有此值则读取最后一行的范围，此值并不能完全反映工作表的有效行数
     *
     * @return 当前工作表的有效区域
     */
    Dimension getDimension();

    /**
     * 判断当前工作表是否隐藏
     *
     * @return {@code true} 当前为“隐藏”工作表
     */
    boolean isHidden();

    /**
     * 判断当前工作表是否显示
     *
     * @return {@code true} 当前为“显示”工作表
     */
    default boolean isShow() {
        return !isHidden();
    }

    /**
     * 设置工作表的表头行号（从1开始）与Excel看到的行号一致
     *
     * @param fromRowNum 表头行的位置（从1开始）
     * @return 当前工作表
     * @throws IndexOutOfBoundsException 如果{@code fromRow}小于1
     */
    default Sheet header(int fromRowNum) {
        return header(fromRowNum, fromRowNum);
    }

    /**
     * 设置工作表的表头行号
     *
     * <p>注意: 指定标题行号后，行指针将移动到标题范围的下一行. 以下方法 {@link #bind(Class)}, {@link #bind(Class, int)},
     * {@link #bind(Class, int, int)}, {@link #rows()}, {@link #dataRows()}, {@link #iterator()},
     * 和 {@link #dataIterator()} 将受影响.</p>
     *
     * @param fromRowNum 表头行的开始位置（从1开始，包含）
     * @param toRowNum   表头行的结束位置（从1开始，包含）
     * @return 当前工作表
     * @throws IndexOutOfBoundsException 如果{@code fromRow}小于1
     * @throws IllegalArgumentException  如果{@code toRow}小于{@code fromRow}
     */
    Sheet header(int fromRowNum, int toRowNum);

    /**
     * 获取当前工作表表头，返回{@link #header}方法设置表头，未指定表头位置时默认取第一个非空行做为表头，如果为多行表头则将使用{@code ':'}拼接
     *
     * <blockquote><pre>
     * +-----------------------------+
     * |       |        收件人        |
     * | 订单号 +-------+------+------+
     * |       |   省  |  市  |   区  |
     * +-------+-------+------+------+
     * </pre></blockquote>
     *
     * <p>以上表头将返回 "订单号", "收件人:省", "收件人:市", "收件人:区"</p>
     *
     * @return 表头行
     */
    Row getHeader();

    /**
     * 绑定数据类型，后续可以通过{@link Row#get}方法直接将行数据转为指定的对象
     *
     * @param clazz 行数据需要转换的对象类型
     * @return 当前工作表
     */
    Sheet bind(Class<?> clazz);

    /**
     * 绑定数据类型并指定表头行号，后续可以通过{@link Row#get}方法直接将行数据转为指定的对象
     *
     * @param clazz      行数据需要转换的对象类型
     * @param fromRowNum 表头行的位置（从1开始）
     * @return 当前工作表
     */
    default Sheet bind(Class<?> clazz, int fromRowNum) {
        return bind(clazz, header(fromRowNum).getHeader());
    }

    /**
     * 绑定数据类型并指定表头行号，后续可以通过{@link Row#get}方法直接将行数据转为指定的对象
     *
     * @param clazz      行数据需要转换的对象类型
     * @param fromRowNum 表头行的位置（从1开始）
     * @param toRowNum   表头行的结束位置（从1开始，包含）
     * @return 当前工作表
     */
    default Sheet bind(Class<?> clazz, int fromRowNum, int toRowNum) {
        return bind(clazz, header(fromRowNum, toRowNum).getHeader());
    }

    /**
     * 绑定数据类型并指定表头，后续可以通过{@link Row#get}方法直接将行数据转为指定的对象
     *
     * @param clazz 行数据需要转换的对象类型
     * @param row   自定义表头
     * @return 当前工作表
     */
    Sheet bind(Class<?> clazz, Row row);

    /**
     * 加载工作表，读取工作表之前必须先使用此方法加载，使用Reader的场景已默认加载无需手动加载
     *
     * @return 当前工作表
     * @throws IOException 读取异常
     */
    Sheet load() throws IOException;

    /**
     * 构建一个行迭代器（包含空行），注意返回的{@code Row}对象是内存共享的所以不能直接收集，
     * 收集数据前需要使用{@link Row#to}方法转为对象或者使用{@link Row#toMap}方法转为Map再收集。
     *
     * @return 行迭代器
     */
    Iterator<Row> iterator();

    /**
     * 构建一个行迭代器（不包含空行），注意返回的{@code Row}对象是内存共享的所以不能直接收集，
     * 收集数据前需要使用{@link Row#to}方法转为对象或者使用{@link Row#toMap}方法转为Map再收集。
     *
     * @return 行迭代器
     */
    Iterator<Row> dataIterator();

    /**
     * 获取当前工作表包含的所有图片
     *
     * @return 图片列表，如果不包含图片则返回{@code null}
     */
    List<Drawings.Picture> listPictures();

    /**
     * 重置游标以重头开始读，可以起到重复读的用处，不过此方法不是必要的，也可以直接通过reader获取对应工作表也可以
     *
     * @return 当前工作表
     * @throws ExcelReadException            读取异常
     * @throws UnsupportedOperationException 如果实现类不支持重复读时抛此异常
     */
    default Sheet reset() {
        throw new UnsupportedOperationException();
    }

    /**
     * 返回一个行流，它与{@link #iterator()}具有相同的功能
     *
     * @return 行流
     */
    default Stream<Row> rows() {
        return StreamSupport.stream(Spliterators.spliteratorUnknownSize(
            iterator(), Spliterator.ORDERED | Spliterator.NONNULL), false);
    }

    /**
     * 返回一个非空行流，它与{@link #dataIterator}具有相同的功能
     *
     * @return 非空行流
     */
    default Stream<Row> dataRows() {
        return StreamSupport.stream(Spliterators.spliteratorUnknownSize(
            dataIterator(), Spliterator.ORDERED | Spliterator.NONNULL), false);
    }


    /**
     * 将当前工作表另存为{@code CSV}格式并保存到{@code path}文件中，默认以{@code UTF-8}字符集保存
     *
     * @param path 另存为CSV文件路径
     * @throws IOException 读写异常
     */
    default void saveAsCSV(Path path) throws IOException {
        saveAsCSV(path, StandardCharsets.UTF_8);
    }

    /**
     * 将当前工作表另存为{@code CSV}格式以{@code UTF-8}字符集保存保存到{@code path}文件中
     *
     * @param path    另存为CSV文件路径
     * @param charset 指定字符集
     * @throws IOException 读写异常
     */
    default void saveAsCSV(Path path, Charset charset) throws IOException {
        Path outPath = FileUtil.getTargetPath(path, getName(), Const.Suffix.CSV), parent = outPath.getParent();
        if (parent != null && !Files.exists(parent)) FileUtil.mkdir(parent);

        saveAsCSV(Files.newOutputStream(outPath), charset);
    }

    /**
     * 将当前工作表另存为{@code CSV}格式并输出到指定字节流
     *
     * @param os 输出字节流
     * @throws IOException 读写异常
     */
    default void saveAsCSV(OutputStream os) throws IOException {
        saveAsCSV(os, StandardCharsets.UTF_8);
    }

    /**
     * 将当前工作表另存为{@code CSV}格式并指定字符集然后输出到指定流
     *
     * @param os      输出流
     * @param charset 字符集
     * @throws IOException 读写异常
     */
    default void saveAsCSV(OutputStream os, Charset charset) throws IOException {
        saveAsCSV(new BufferedWriter(new OutputStreamWriter(os, charset)));
    }

    /**
     * 将当前工作表另存为{@code CSV}格式并输出到指定流
     *
     * @param bw 输出流
     * @throws IOException 读写异常
     */
    default void saveAsCSV(BufferedWriter bw) throws IOException {
        try (CSVUtil.Writer writer = CSVUtil.newWriter(bw)) {
            int rowNum = 1;
            for (Iterator<Row> iter = iterator(); iter.hasNext(); ) {
                Row row = iter.next();
                // 保持与xlsx相同行号
                if (row.getRowNum() - rowNum > 1) {
                    for (; ++rowNum < row.getRowNum(); writer.newLine());
                } else rowNum = row.getRowNum();
                if (row.isBlank()) {
                    writer.newLine();
                    continue;
                }
                for (int i = 0; i < row.lc; i++) {
                    Cell c = row.cells[i];
                    switch (c.t) {
                        case SST:
                            if (c.stringVal == null) c.setString(row.sst.get(c.intVal));
                        case INLINESTR:
                        case FUNCTION:
                            writer.write(c.stringVal);
                            break;
                        case NUMERIC:
                            if (!row.styles.isDate(c.xf)) writer.write(c.intVal);
                            else if (c.intVal > 0) writer.write(toLocalDate(c.intVal).toString());
                            // 时分秒为00:00:00时读取为整数0
                            else writer.write(toTimeChars(LocalTime.MIN));
                            break;
                        case LONG:
                            writer.write(c.longVal);
                            break;
                        case DECIMAL:
                            if (!row.styles.isDate(c.xf)) writer.write(c.decimal.toString());
                            else if (c.decimal.compareTo(BigDecimal.ONE) >= 0) writer.write(toDateTimeString(toTimestamp(c.decimal.doubleValue())));
                            else writer.write(toTimeChars(toLocalTime(c.decimal.doubleValue())));
                            break;
                        case DOUBLE:
                            if (!row.styles.isDate(c.xf)) writer.write(c.doubleVal);
                            else if (c.doubleVal >= 1.0D) writer.write(toDateTimeString(toTimestamp(c.doubleVal)));
                            else writer.write(toTimeChars(toLocalTime(c.doubleVal)));
                            break;
                        case BOOL:
                            writer.write(c.boolVal);
                            break;
                        default:
                            writer.writeEmpty();
                    }
                }
                writer.newLine();
            }
        }
    }

    /**
     * 强制匹配，即使没有{@link org.ttzero.excel.annotation.ExcelColumn}注解的字段也会强制匹配
     *
     * <p>将行数据转换为Java对象时默认情况下只匹配带有ExcelColumn注释的字段。
     * {@code forceImport}将跳过此限制，有&#40;ExcelColumn注释的依然按注解匹配，
     * 没有&#40;ExcelColumn注释的字段将与字段名匹配</p>
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
     * <pre>reader.sheet(0).setHeaderColumnReadOption(HeaderRow.FORCE_IMPORT | HeaderRow.IGNORE_CASE)</pre>
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

//    /**
//     * 将工作表转为{@code CalcSheet}以解析单元格公式
//     *
//     * @return {@link CalcSheet}
//     * @deprecated 使用 {@link #asFullSheet()} 替换，{@code FullSheet}包含{@code CalcSheet}所有功能
//     */
//    @Deprecated
//    CalcSheet asCalcSheet();
//
//    /**
//     * 将工作表转为{@code MergeSheet}，它将复制合并单元格的首坐标值到合并范围内的其它单元格中
//     *
//     * @return {@link MergeSheet}
//     * @deprecated 使用 {@link #asFullSheet()} 替换，{@code FullSheet}包含{@code MergeSheet}所有功能
//     */
//    @Deprecated
//    MergeSheet asMergeSheet();

    /**
     * 将工作表转为{@code FullSheet}支持全属性读取
     *
     * @return {@link FullSheet}
     */
    FullSheet asFullSheet();
}
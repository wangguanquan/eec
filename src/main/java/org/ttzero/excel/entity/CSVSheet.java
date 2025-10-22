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

package org.ttzero.excel.entity;

import org.ttzero.excel.manager.Const;
import org.ttzero.excel.reader.Cell;
import org.ttzero.excel.util.CSVUtil;
import org.ttzero.excel.util.FileUtil;

import java.io.BufferedWriter;
import java.io.IOException;
import java.io.InputStream;
import java.io.Reader;
import java.nio.charset.Charset;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;

import static org.ttzero.excel.util.FileUtil.exists;

/**
 * {@code CSVSheet}的数据源为csv文件，可用于将csv文件转为Excel工作表，
 * 通过{@link CSVUtil}提供的迭代器使得{@code CSVSheet}与{@code ListSheet}工作表具有
 * 相似的切片属性，输出协议调用{@code nextBlock}获取分片数据时{@code CSVSheet}从CSVIterator
 * 中逐行读取数据并输出以此控制整个过程对内存的消耗
 *
 * <p>默认情况下CSV文件第一行数据将做为Excel表头，如果要忽略这个设定则需要调用{@link #ignoreHeader()}方法忽略表头，
 * 忽略表头后第一行数据将被视为普通数据导了，也就是说无论是否忽略表头CSV的数据都将被完全复制到Excel文件中，
 * 只是默认第一行将添加表头样式</p>
 *
 * @author guanquan.wang at 2019-09-26 08:33
 */
public class CSVSheet extends Sheet {

    /**
     * csv文件路径
     */
    protected Path path;
    /**
     * csv Reader
     */
    protected CSVUtil.Reader reader;
    /**
     * csv行迭代器，配合工作表输出协议获取数据可以极大降低内存消耗
     */
    protected CSVUtil.RowsIterator iterator;
    /**
     * 是否需要清理临时资源，实例化时如果传入{@code InputStream}或{@code Reader}时会先将数据保存到临时文件
     * 然后创建迭代器逐行读取数据，这个过程产生的临时文件会在关闭工作表时被一起清理
     */
    protected boolean shouldClean;
    /**
     * 指定读取CSV使用的字符集
     */
    protected Charset charset;
    /**
     * Delimiter char
     */
    protected char delimiter = ',';

    /**
     * 实例化工作表，未指定工作表名称时默认以{@code 'Sheet'+id}命名
     */
    public CSVSheet() {
        super();
    }

    /**
     * 实例化工作表并指定csv文件路径
     *
     * @param path csv文件路径
     */
    public CSVSheet(Path path) {
        this.path = path;
    }

    /**
     * 实例化工作表并指定工作表名和csv文件路径
     *
     * @param name 工作表名
     * @param path csv文件路径
     */
    public CSVSheet(String name, Path path) {
        super(name);
        this.path = path;
    }

    /**
     * 实例化工作表并指定csv文件字节流
     *
     * @param is csv文件字节流
     * @throws IOException if I/O error occur.
     */
    public CSVSheet(InputStream is) throws IOException {
        this(null, is);
    }

    /**
     * 实例化工作表并指定工作表名和csv文件字节流
     *
     * @param name 工作表名
     * @param is   csv文件字节流
     * @throws IOException if I/O error occur.
     */
    public CSVSheet(String name, InputStream is) throws IOException {
        super(name);
        path = Files.createTempFile(Const.EEC_PREFIX, String.valueOf(id));
        Files.copy(is, path, StandardCopyOption.REPLACE_EXISTING);
        shouldClean = true;
    }

    /**
     * 实例化工作表并指定csv文件字符流
     *
     * @param reader csv文件字符流
     * @throws IOException if I/O error occur.
     */
    public CSVSheet(Reader reader) throws IOException {
        this(null, reader);
    }

    /**
     * 实例化工作表并指定工作表名和csv文件字符流
     *
     * @param name   工作表名
     * @param reader csv文件字符流
     * @throws IOException if I/O error occur.
     */
    public CSVSheet(String name, Reader reader) throws IOException {
        super(name);
        path = Files.createTempFile(Const.EEC_PREFIX, String.valueOf(id));
        shouldClean = true;
        char[] chars = new char[8192];
        int n;
        try (BufferedWriter bw = Files.newBufferedWriter(path)) {
            while ((n = reader.read(chars)) > 0) {
                bw.write(chars, 0, n);
            }
        }
    }

    /**
     * 清理临时文件
     *
     * @throws IOException if I/O error occur
     */
    @Override
    public void close() throws IOException {
        // 最后一个Sheet关闭CSV流
        if (shouldClose) {
            if (iterator != null) iterator.close();
            if (reader != null) reader.close();
            if (shouldClean) FileUtil.rm_rf(path);
        }
        super.close();
    }

    // Create CSV iterator
    private void init() throws IOException {
        assert path != null && exists(path);
        reader = CSVUtil.newReader(path, delimiter, charset);
        iterator = reader.sharedIterator();
    }

    /**
     * 重置{@code RowBlock}行块数据，使用csv行迭代器逐行读取数据并重置行块，由于csv格式并不包含任务样式
     * 所以{@code CSVSheet}并不支持任务样式设定
     */
    @Override
    protected void resetBlockData() {
        int n = 0, limit = getRowLimit();
        boolean hasNext = true;
        for (int rbs = rowBlock.capacity(); n++ < rbs && rows < limit && (hasNext = iterator.hasNext()); rows++) {
            Row row = rowBlock.next();
            row.index = rows;
            row.height = getRowHeight();
            resetRowData(row, iterator.next());
        }

        // Paging
        if (rows >= limit) {
            shouldClose = false;
            rowBlock.markEOF();
            CSVSheet copy = getClass().cast(clone());
            copy.shouldClose = true;
            workbook.insertSheet(id, copy);
        } else if (!hasNext) rowBlock.markEOF();
    }

    @Override
    protected Column[] getHeaderColumns() {
        if (headerReady) return columns;
        try {
            // Create CSV iterator
            init();
            if (!iterator.hasNext()) return columns;
            String[] rows = iterator.next();
            columns = new Column[rows.length];
            for (int i = 0; i < rows.length; i++) {
                // FIXME the column type
                columns[i] = new Column(getNonHeader() != 1 ? rows[i] : null, String.class);
                columns[i].styles = workbook.getStyles();
            }
        } catch (IOException e) {
            throw new ExcelWriteException(e);
        }
        return columns;
    }

    @Override
    public void checkColumnLimit() {
        super.checkColumnLimit();
        if (getNonHeader() == 1) {
            ((CSVUtil.SharedRowsIterator) iterator).retain();
        }
    }

    @Override
    protected void mergeHeaderCellsIfEquals() { }

    /**
     * 设置读取CSV使用的字符集
     *
     * @param charset 指定字符集
     * @return 当前工作表
     */
    public CSVSheet setCharset(Charset charset) {
        this.charset = charset;
        return this;
    }

    /**
     * Settings delimiter char
     *
     * @param delimiter delimiter char
     * @return current CSVSheet
     */
    public CSVSheet setDelimiter(char delimiter) {
        this.delimiter = delimiter;
        return this;
    }

    /**
     * 重置单行数据
     *
     * @param row Excel行
     * @param rowData 行数据
     */
    protected void resetRowData(Row row, String[] rowData) {
        int len = Math.min(columns.length, rowData.length);
        Cell[] cells = row.realloc(len);
        for (int i = 0; i < len; i++) {
            resetCellValueAndStyle(row, cells[i], rowData, rowData[i], columns[i]);
        }
    }

    /**
     * 重置单元格数据和样式
     *
     * @param row Excel行
     * @param cell Excel单元格
     * @param rowData 行数据
     * @param cellData 单元格数据
     * @param column 单列表头
     */
    protected void resetCellValueAndStyle(Row row, Cell cell, String[] rowData, String cellData, Column column) {
        cell.setString(cellData);
        cell.xf = cellValueAndStyle.getStyleIndex(row, column, cellData);
    }
}

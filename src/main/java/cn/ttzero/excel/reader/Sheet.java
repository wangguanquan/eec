/*
 * Copyright (c) 2009, guanquan.wang@yandex.com All Rights Reserved.
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

package cn.ttzero.excel.reader;

import cn.ttzero.excel.entity.TooManyColumnsException;
import cn.ttzero.excel.manager.Const;
import cn.ttzero.excel.util.StringUtil;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.UncheckedIOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Stream;
import java.util.stream.StreamSupport;

/**
 * 对应Excel文件各Sheet页，包含隐藏Sheet页
 *
 * Create by guanquan.wang at 2018-09-22
 */
public class Sheet implements AutoCloseable {
    private Logger logger = LogManager.getLogger(getClass());
    Sheet() {}

    private String name;
    private int index // per sheet index of workbook
            , size = -1; // size of rows per sheet
    private Path path;
    private SharedString sst;
    private int startRow = -1; // row index of data
    private Row header;
    private boolean hidden; // state hidden

    void setName(String name) {
        this.name = name;
    }

    void setPath(Path path) {
        this.path = path;
    }

    void setSst(SharedString sst) {
        this.sst = sst;
    }

    /**
     * @return sheet名
     */
    public String getName() {
        return name;
    }

    /**
     * @return sheet位于workbook的位置
     */
    public int getIndex() {
        return index;
    }

    void setIndex(int index) {
        this.index = index;
    }

    /**
     * size of sheet. -1: unknown size
     * @return number of rows
     */
    public int getSize() {
        return size;
    }

    public int getStartRow() {
        return startRow;
    }

    /**
     * Test Worksheet is hidden
     */
    public boolean isHidden() {
        return hidden;
    }

    /**
     * Test Worksheet is show
     */
    public boolean isShow() {
        return !hidden;
    }

    /**
     * Set Worksheet state
     */
    Sheet setHidden(boolean hidden) {
        this.hidden = hidden;
        return this;
    }

    /**
     * 返回列表头，第一个非空行默认为头部信息
     * @return HeaderRow
     */
    public Row getHeader() {
        if (header == null && !heof) {
            Row row = findRow0();
            if (row != null) {
                header = row.asHeader();
                sRow.setHr((Row.HeaderRow) header);
            }
        }
        return header;
    }

    /**
     * 绑定对象
     * @param clazz 对象类型
     * @return sheet
     */
    public Sheet bind(Class<?> clazz) {
        if (getHeader() != null) {
            try {
                ((Row.HeaderRow) header).setClassOnce(clazz);
            } catch (IllegalAccessException | InstantiationException e) {
                throw new ExcelReadException(e);
            }
        }
        return this;
    }

    @Override
    public String toString() {
        return name;
    }

    /////////////////////////////////Read sheet file/////////////////////////////////
    private BufferedReader reader;
    private char[] cb; // buffer
    private int nChar, length;
    private boolean eof = false, heof = false; // OPTIONS = false

    private Row sRow;

    /**
     * load sheet.xml as BufferedReader
     * @return Sheet
     * @throws IOException
     */
    Sheet load() throws IOException {
        logger.debug("load {}", path.toString());
        reader = Files.newBufferedReader(path);
        cb = new char[8192];
        nChar = 0;
        loopA: for ( ; ; ) {
            length = reader.read(cb);
            if (length < 11) break;
            // read size
            if (nChar == 0) {
                String line = new String(cb, 56, 1024);
                String size = "<dimension ref=\"";
                int index = line.indexOf(size), end = index > 0 ? line.indexOf('"', index+=size.length()) : -1;
                if (end > 0) {
                    String l_ = line.substring(index, end);
                    Pattern pat = Pattern.compile("([A-Z]+)(\\d+):([A-Z]+)(\\d+)");
                    Matcher mat = pat.matcher(l_);
                    if (mat.matches()) {
                        int from = Integer.parseInt(mat.group(2)), to = Integer.parseInt(mat.group(4));
                        this.startRow = from;
                        this.size = to - from + 1;
                        int c1 = col2Int(mat.group(1)), c2 = col2Int(mat.group(3)), columnSize = c2 - c1 + 1;
                        if (columnSize > Const.Limit.MAX_COLUMNS_ON_SHEET) {
                            throw new TooManyColumnsException(columnSize);
                        }
                        logger.debug("column size: {}", columnSize);
//                        OPTIONS = columnSize > 128; // size more than DX
                    } else {
                        pat = Pattern.compile("[A-Z]+(\\d+)");
                        mat = pat.matcher(l_);
                        if (mat.matches()) {
                            this.startRow = Integer.parseInt(mat.group(1));
                        }
                    }
                    nChar += end;
                }
            }
            // find index of <sheetData>
            for (; nChar < length - 12; nChar++) {
                if (cb[nChar] == '<' && cb[nChar+1] == 's' && cb[nChar+2] == 'h'
                        && cb[nChar+3] == 'e' && cb[nChar+4] == 'e' && cb[nChar+5] == 't'
                        && cb[nChar+6] == 'D' && cb[nChar+7] == 'a' && cb[nChar+8] == 't'
                        && cb[nChar+9] == 'a' && (cb[nChar+10] == '>' || cb[nChar+10] == '/' && cb[nChar+11] == '>')) {
                    nChar += 11;
                    break loopA;
                }
            }
        }
        // Empty sheet
        if (cb[nChar] == '>') {
            nChar++;
            eof = true;
        } else {
            sRow = new Row(sst, this.startRow > 0 ? this.startRow : 1); // share row space
        }

        return this;
    }

    /**
     * iterator rows
     * @return Row
     * @throws IOException
     */
    private Row nextRow() throws IOException {
        if (eof) return null;
        boolean endTag = false;
        int start = nChar;
        // find end of row tag
        for ( ; ++nChar < length && cb[nChar] != '>'; );
        // Empty Row
        if (cb[nChar++ - 1] == '/') {
            return sRow.empty(cb, start, nChar - start);
        }
        // Not empty
        for (; nChar < length - 6; nChar++) {
            if (cb[nChar] == '<' && cb[nChar+1] == '/' && cb[nChar+2] == 'r'
                    && cb[nChar+3] == 'o' && cb[nChar+4] == 'w' && cb[nChar+5] == '>') {
                nChar += 6;
                endTag = true;
                break;
            }
        }

        /* Load more when not found end of row tag */
        if (!endTag) {
            int n;
            if (start == 0) {
                char[] _cb = new char[cb.length << 1];
                System.arraycopy(cb, start, _cb, 0, n = length - start);
                cb = _cb;
            } else {
                System.arraycopy(cb, start, cb, 0, n = length - start);
            }
            length = reader.read(cb, n, cb.length - n);
            // end of file
            if (length < 0) {
                eof = true;
                reader.close(); // close reader
                reader = null; // wait GC
                logger.debug("end of file.");
                return null;
            }
            nChar = 0;
            length += n;
            return nextRow();
        }

        // share row
        return sRow.with(cb, start, nChar - start);
    }

    protected Row findRow0() {
        char[] cb = new char[8192];
        int nChar = 0, length;
        // reload file
        try (BufferedReader reader = Files.newBufferedReader(path)) {
            loopA: for ( ; ; ) {
                length = reader.read(cb);
                // find index of <sheetData>
                for (; nChar < length - 12; nChar++) {
                    if (cb[nChar] == '<' && cb[nChar + 1] == 's' && cb[nChar + 2] == 'h'
                            && cb[nChar + 3] == 'e' && cb[nChar + 4] == 'e' && cb[nChar + 5] == 't'
                            && cb[nChar + 6] == 'D' && cb[nChar + 7] == 'a' && cb[nChar + 8] == 't'
                            && cb[nChar + 9] == 'a' && (cb[nChar + 10] == '>' || cb[nChar+10] == '/' && cb[nChar+11] == '>')) {
                        nChar += 11;
                        break loopA;
                    }
                }
            }
        } catch (IOException e) {
            logger.error("Read header row error.");
            return null;
        }
        boolean eof = cb[nChar] == '>';
        if (eof) {
            this.heof = eof;
            return null;
        }

        boolean endTag;
        int start;
        // find the first not null row
        loopB: while (true) {
            start = nChar;
            for (; cb[++nChar] != '>' && nChar < length; );
            // Empty Row
            if (cb[nChar++ - 1] != '/') {
                // find end of row tag
                for (; nChar < length - 6; nChar++) {
                    if (cb[nChar] == '<' && cb[nChar + 1] == '/' && cb[nChar + 2] == 'r'
                            && cb[nChar + 3] == 'o' && cb[nChar + 4] == 'w' && cb[nChar + 5] == '>') {
                        nChar += 6;
                        endTag = true;
                        break loopB;
                    }
                }
            }
        }

        if (!endTag) {
            // too big
            return null;
        }

        // row
        return new Row(sst, this.startRow > 0 ? this.startRow : 1).with(cb, start, nChar - start);
    }

    /**
     * 迭代每行数据包含头部信息和空行数据
     * @return 迭代器
     */
    public Iterator<Row> iterator() {
        return iter;
    }

    /**
     * 迭代数据行，不包含头部信息和空行
     * @return 迭代器
     */
    public Iterator<Row> dataIterator() {
        if (nIter.hasNext()) {
            Row row = nIter.next();
            if (header == null) header = row.asHeader();
        }
        return nIter;
    }

    // iterator data rows
    private Iterator<Row> nIter = new Iterator<Row>() {
        Row nextRow = null;

        @Override
        public boolean hasNext() {
            if (nextRow != null) {
                return true;
            } else {
                try {
                    // Skip empty rows
                    for ( ; (nextRow = nextRow()) != null && nextRow.isEmpty(); );
                    return nextRow != null;
                } catch (IOException e) {
                    throw new UncheckedIOException(e);
                }
            }
        }

        @Override
        public Row next() {
            if (nextRow != null || hasNext()) {
                Row next = nextRow;
                nextRow = null;
                return next;
            } else {
                throw new NoSuchElementException();
            }
        }
    };

    // iterator foreach rows
    private Iterator<Row> iter = new Iterator<Row>() {
        Row nextRow = null;

        @Override
        public boolean hasNext() {
            if (nextRow != null) {
                return true;
            } else {
                try {
                    nextRow = nextRow();
                    return (nextRow != null);
                } catch (IOException e) {
                    throw new UncheckedIOException(e);
                }
            }
        }

        @Override
        public Row next() {
            if (nextRow != null || hasNext()) {
                Row next = nextRow;
                nextRow = null;
                return next;
            } else {
                throw new NoSuchElementException();
            }
        }
    };

    /**
     * stream of all rows
     * @return a {@code Stream<Row>} providing the lines of row
     *         described by this {@code Sheet}
     * @since 1.8
     */
    public Stream<Row> rows() {
        return StreamSupport.stream(Spliterators.spliteratorUnknownSize(
                iter, Spliterator.ORDERED | Spliterator.NONNULL), false);
    }

    /**
     * stream with out header row and empty rows
     * @return a {@code Stream<Row>} providing the lines of row
     *         described by this {@code Sheet}
     * @since 1.8
     */
    public Stream<Row> dataRows() {
        if (nIter.hasNext()) {
            Row row = nIter.next();
            if (header == null) header = row.asHeader();
        }
        return StreamSupport.stream(Spliterators.spliteratorUnknownSize(
                nIter, Spliterator.ORDERED | Spliterator.NONNULL), false);
    }

    /**
     * column mark to int
     * @param col column mark
     * @return int value
     */
    private int col2Int(String col) {
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

    /**
     * close reader
     * @throws IOException
     */
    public void close() throws IOException {
        cb = null;
        if (reader != null) {
            reader.close();
        }
        if (sst != null)
            sst.close();
    }
}

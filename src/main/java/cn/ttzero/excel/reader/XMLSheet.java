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

package cn.ttzero.excel.reader;

import cn.ttzero.excel.entity.TooManyColumnsException;
import cn.ttzero.excel.manager.Const;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import java.io.BufferedReader;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * 对应Excel文件各Sheet页，包含隐藏Sheet页
 * <p>
 * Create by guanquan.wang on 2018-09-22
 */
class XMLSheet implements Sheet {
    private Logger logger = LogManager.getLogger(getClass());

    XMLSheet() { }

    private String name;
    private int index // per sheet index of workbook
        , size = -1; // size of rows per sheet
    private Path path;
    private SharedStrings sst;
    private int startRow = -1; // row index of data
    private HeaderRow header;
    private boolean hidden; // state hidden

    void setName(String name) {
        this.name = name;
    }

    void setPath(Path path) {
        this.path = path;
    }

    void setSst(SharedStrings sst) {
        this.sst = sst;
    }

    /**
     * The worksheet name
     *
     * @return the sheet name
     */
    @Override
    public String getName() {
        return name;
    }

    /**
     * The index of worksheet located at the workbook
     *
     * @return the index(zero base)
     */
    @Override
    public int getIndex() {
        return index;
    }

    void setIndex(int index) {
        this.index = index;
    }

    /**
     * size of rows.
     *
     * @return size of rows
     * -1: unknown size
     */
    @Override
    public int getSize() {
        return size;
    }

    /**
     * The index of first used row
     *
     * @return the index
     */
    public int getFirstRow() {
        return startRow;
    }

    /**
     * Test Worksheet is hidden
     */
    @Override
    public boolean isHidden() {
        return hidden;
    }

    /**
     * Set Worksheet state
     */
    XMLSheet setHidden(boolean hidden) {
        this.hidden = hidden;
        return this;
    }

    /**
     * Returns the header of the list.
     * The first non-empty line defaults to the header information.
     *
     * @return the HeaderRow
     */
    @Override
    public Row getHeader() {
        if (header == null && !heof) {
            XMLRow row = findRow0();
            if (row != null) {
                header = row.asHeader();
                sRow.setHr(header);
            }
        }
        return header;
    }

    /**
     * Set the binding type
     *
     * @param clazz the binding type
     * @return sheet
     */
    @Override
    public XMLSheet bind(Class<?> clazz) {
        if (getHeader() != null) {
            try {
                header.setClassOnce(clazz);
            } catch (IllegalAccessException | InstantiationException e) {
                throw new ExcelReadException(e);
            }
        }
        return this;
    }

    @Override
    public String toString() {
        return "Sheet name: " + name + " has " + size + " rows.";
    }

    /////////////////////////////////Read sheet file/////////////////////////////////
    private BufferedReader reader;
    private char[] cb; // buffer
    private int nChar, length;
    private boolean eof = false, heof = false; // OPTIONS = false

    private XMLRow sRow;

    /**
     * Load sheet.xml as BufferedReader
     *
     * @return Sheet
     * @throws IOException if io error occur
     */
    public XMLSheet load() throws IOException {
        logger.debug("load {}", path.toString());
        reader = Files.newBufferedReader(path);
        cb = new char[8192];
        nChar = 0;
        loopA: for (; ; ) {
            length = reader.read(cb);
            if (length < 11) break;
            // read size
            if (nChar == 0) {
                String line = new String(cb, 56, 1024);
                String size = "<dimension ref=\"";
                int index = line.indexOf(size), end = index > 0 ? line.indexOf('"', index += size.length()) : -1;
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
                            throw new TooManyColumnsException(columnSize, Const.Limit.MAX_COLUMNS_ON_SHEET);
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
                if (cb[nChar] == '<' && cb[nChar + 1] == 's' && cb[nChar + 2] == 'h'
                    && cb[nChar + 3] == 'e' && cb[nChar + 4] == 'e' && cb[nChar + 5] == 't'
                    && cb[nChar + 6] == 'D' && cb[nChar + 7] == 'a' && cb[nChar + 8] == 't'
                    && cb[nChar + 9] == 'a' && (cb[nChar + 10] == '>' || cb[nChar + 10] == '/' && cb[nChar + 11] == '>')) {
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
            eof = false;
            sRow = new XMLRow(sst, this.startRow > 0 ? this.startRow : 1); // share row space
        }

        return this;
    }

    /**
     * iterator rows
     *
     * @return Row
     * @throws IOException if io error occur
     */
    private XMLRow nextRow() throws IOException {
        if (eof) return null;
        boolean endTag = false;
        int start = nChar;
        // find end of row tag
        for (; ++nChar < length && cb[nChar] != '>'; ) ;
        // Empty Row
        if (cb[nChar++ - 1] == '/') {
            return sRow.empty(cb, start, nChar - start);
        }
        // Not empty
        for (; nChar < length - 6; nChar++) {
            if (cb[nChar] == '<' && cb[nChar + 1] == '/' && cb[nChar + 2] == 'r'
                && cb[nChar + 3] == 'o' && cb[nChar + 4] == 'w' && cb[nChar + 5] == '>') {
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

    protected XMLRow findRow0() {
        char[] cb = new char[8192];
        int nChar = 0, length;
        // reload file
        try (BufferedReader reader = Files.newBufferedReader(path)) {
            loopA: for (; ; ) {
                length = reader.read(cb);
                // find index of <sheetData>
                for (; nChar < length - 12; nChar++) {
                    if (cb[nChar] == '<' && cb[nChar + 1] == 's' && cb[nChar + 2] == 'h'
                        && cb[nChar + 3] == 'e' && cb[nChar + 4] == 'e' && cb[nChar + 5] == 't'
                        && cb[nChar + 6] == 'D' && cb[nChar + 7] == 'a' && cb[nChar + 8] == 't'
                        && cb[nChar + 9] == 'a' && (cb[nChar + 10] == '>' || cb[nChar + 10] == '/' && cb[nChar + 11] == '>')) {
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
            for (; cb[++nChar] != '>' && nChar < length; ) ;
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
        return new XMLRow(sst, this.startRow > 0 ? this.startRow : 1).with(cb, start, nChar - start);
    }

    /**
     * Iterating each row of data contains header information and blank lines
     *
     * @return a row iterator
     */
    @Override
    public Iterator<Row> iterator() {
        return new RowSetIterator(this::nextRow, false);
    }

    /**
     * Iterating over data rows without header information and blank lines
     *
     * @return a row iterator
     */
    @Override
    public Iterator<Row> dataIterator() {
        // iterator data rows
        Iterator<Row> nIter = new RowSetIterator(this::nextRow, true);
        if (nIter.hasNext()) {
            Row row = nIter.next();
            if (header == null)
                header = row.asHeader();
        }
        return nIter;
    }

    /**
     * close reader
     *
     * @throws IOException if io error occur
     */
    @Override
    public void close() throws IOException {
        cb = null;
        if (reader != null) {
            reader.close();
        }
    }

    /**
     * Reset the {@link XMLSheet}'s row index to begging
     *
     * @return the unread {@link XMLSheet}
     */
    @Override
    public XMLSheet reset() throws IOException {
        // Close the opening reader
        close();
        // Reload
        return load();
    }
}

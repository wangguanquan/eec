/*
 * Copyright (c) 2019-2021, guanquan.wang@yandex.com All Rights Reserved.
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

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.ttzero.excel.entity.TooManyColumnsException;
import org.ttzero.excel.entity.style.Styles;
import org.ttzero.excel.manager.Const;

import java.io.BufferedReader;
import java.io.IOException;
import java.nio.ByteBuffer;
import java.nio.CharBuffer;
import java.nio.channels.SeekableByteChannel;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardOpenOption;
import java.util.Arrays;
import java.util.Iterator;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * The open-xml format Worksheet
 * <p>
 * Create by guanquan.wang on 2018-09-22
 */
class XMLSheet implements Sheet {
    private Logger logger = LogManager.getLogger(getClass());

    XMLSheet() { }

    private String name;
    private int index; // per sheet index of workbook
//        , size = -1; // size of rows per sheet
    private Path path;
    /**
     * The Shared String Table
     */
    private SharedStrings sst;
    /**
     * The {@link Styles}
     */
    private Styles styles;
    private int startRow = -1; // row index of data
    private HeaderRow header;
    private boolean hidden; // state hidden
    /*
     The range address of the used area in the current sheet
     Offset | Size | Contents
     -------|------|----------
        0   |   4  | Index to first used row
        4   |   4  | Index to last used row, increased by 1
        8   |   2  | Index to first used column
       10   |   2  | Index to last used column, increased by 1
     */
    private long dimension;
    private int rc; // Range of column
    private long[] calc; // Array of formula
    private boolean nonCalc = true;

    /**
     * Setting the worksheet name
     *
     * @param name the worksheet name
     */
    void setName(String name) {
        this.name = name;
    }

    /**
     * Setting the worksheet xml path
     *
     * @param path the temp path
     */
    void setPath(Path path) {
        this.path = path;
    }

    /**
     * Setting the Shared String Table
     *
     * @param sst the {@link SharedStrings}
     */
    void setSst(SharedStrings sst) {
        this.sst = sst;
    }

    /**
     * Setting {@link Styles}
     *
     * @param styles the {@link Styles}
     */
    void setStyles(Styles styles) {
        this.styles = styles;
    }

    /**
     * Setting formula array
     *
     * @param calc array of formula
     */
    void setCalc(long[] calc) {
        this.calc = calc;
        this.nonCalc = calc == null || calc.length == 0;
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
        if (dimension == 0) return -1;
        return (int) ((dimension >>> 32) - (dimension & 0x7FFFFFFF));
    }

    /**
     * Returns The range address of the used area in
     * the current sheet
     * <p>
     * NOTE: This method can only guarantee accurate row ranges
     *
     * @return worksheet {@link Dimension} ranges
     */
    @Override
    public Dimension getDimension() {
        // Read last rows info
        if (dimension == 0 || ((dimension >> 32) & 0x7FFFFFFF) == 0) {
            parseDimension();
        }
        return new Dimension((int)(dimension & 0x7FFFFFFF)
            , (int) (dimension >>> 32)
            , (short) (rc & 0x7FFF)
            , (short) (rc >> 16));
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
        return "Sheet name: " + name + " has " + getDimension();
    }

    /////////////////////////////////Read sheet file/////////////////////////////////
    private BufferedReader reader;
    private char[] cb; // buffer
    private int nChar, length;
    private boolean eof = false, heof = false; // OPTIONS = false
    private long mark;

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
                        dimension = ((long) to << 32) | from;
                        this.startRow = from;
//                        this.size = to - from + 1;
                        int c1 = col2Int(mat.group(1)), c2 = col2Int(mat.group(3)), columnSize = c2 - c1 + 1;
                        if (columnSize > Const.Limit.MAX_COLUMNS_ON_SHEET) {
                            throw new TooManyColumnsException(columnSize, Const.Limit.MAX_COLUMNS_ON_SHEET);
                        }
                        rc = (c2 << 16) | c1 & 0x7FFF;
                        logger.debug("Dimension-Range: {\"first-row\": {}, \"last-row\": {}" +
                            ", \"first-column\": {}, \"last-column\": {}}", from, to, c1, c2);
//                        OPTIONS = columnSize > 128; // size more than DX
                    } else {
                        pat = Pattern.compile("([A-Z]+)(\\d+)");
                        mat = pat.matcher(l_);
                        if (mat.matches()) {
                            this.startRow = Integer.parseInt(mat.group(2));
                            dimension = this.startRow;
                            rc = col2Int(mat.group(1));
                            logger.debug("Dimension-Range: {\"first-row\": {}" +
                                ", \"first-column\": {}}", this.startRow, rc);
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
            sRow = new XMLRow(sst, styles, this.startRow > 0 ? this.startRow : 1, this::findCalc); // share row space
        }

        mark = nChar;
        return this;
    }

    /**
     * iterator rows
     *
     * @return Row
     */
    private XMLRow nextRow() {
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
            try {
                length = reader.read(cb, n, cb.length - n);
                // end of file
                if (length < 0) {
                    eof = true;
                    reader.close(); // close reader
                    reader = null; // wait GC
                    logger.debug("end of file.");
                    return null;
                }
            } catch (IOException e) {
                throw new ExcelReadException("Parse row data error", e);
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
            this.heof = true;
            return null;
        }

        boolean endTag = false;
        int start;
        // find the first not null row
        loopB: while (true) {
            start = nChar;
            for (; cb[++nChar] != '>' && nChar < length; ) ;
            if (nChar >= length - 6) break;
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
        return new XMLRow(sst, styles, this.startRow > 0 ? this.startRow : 1, this::findCalc).with(cb, start, nChar - start);
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
    public XMLSheet reset() {
        try {
            // Close the opening reader
            if (reader != null) {
                reader.close();
            }
            // Reload
            reader = Files.newBufferedReader(path);
            reader.skip(mark);
            length = reader.read(cb);
            nChar = 0;
            eof = length <= 0;
        } catch (IOException e) {
            throw new ExcelReadException("Reset worksheet[" + getName() + "] error occur.", e);
        }

        return this;
    }

    /*
    If the Dimension information is missing from the header,
    you can skip to the end of the file and look at the line
    number of the last line to confirm the scope of the entire worksheet.
     */
    private void parseDimension() {
        try (SeekableByteChannel channel = Files.newByteChannel(path, StandardOpenOption.READ)) {
            long fileSize = Files.size(path);
            int block = (int) Math.min(1 << 11, fileSize);
            ByteBuffer buffer = ByteBuffer.allocate(block);
            byte[] left = null;
            int left_size = 0;
            for (; ;) {
                channel.position(fileSize - block + left_size);
                channel.read(buffer);
                fileSize -= buffer.limit();
                if (left_size > 0) {
                    buffer.limit(block);
                    buffer.put(left, 0, left_size);
                    left_size = 0;
                }
                buffer.flip();

                CharBuffer charBuffer = StandardCharsets.UTF_8.decode(buffer);
                int limit = charBuffer.limit(), i = limit - 1;
                // <row r="
                for (; i >= 7 && (charBuffer.get(i) != '"' || charBuffer.get(i - 1) != '='
                    || charBuffer.get(i - 2) != 'r' || charBuffer.get(i - 3) > ' '
                    || charBuffer.get(i - 4) != 'w' || charBuffer.get(i - 5) != 'o'
                    || charBuffer.get(i - 6) != 'r' || charBuffer.get(i - 7) != '<'); i--);

                // Not Found
                if (i < 7) {
                    for (; i < limit && charBuffer.get(i) != '>'; i++);
                    i++;
                    if (i < limit - 1) {
                        charBuffer.position(i);
                        int newLimit = StandardCharsets.UTF_8.encode(charBuffer).limit();
                        int last_size = buffer.limit() - newLimit;
                        if (left == null || last_size > left.length) {
                            left = new byte[last_size];
                        }
                        buffer.position(0);
                        buffer.get(left, 0, last_size);
                        left_size = last_size;
                    }
                }
                // Found the last row
                else {
                    int[] info = innerParse(charBuffer, ++i);
                    rc = (info[2] << 16) | (info[1] & 0x7FFF);
                    dimension = ((long) info[0]) << 32;
                    break;
                }
                buffer.position(0);
                buffer.limit(buffer.capacity() - left_size);
            }
            // Skip if the first row is known
            if ((dimension & 0x7FFFFFFF) > 0L) return;
            if (mark > 0) {
                channel.position(mark);
                buffer.clear();
                channel.read(buffer);
                buffer.flip();

                CharBuffer charBuffer = StandardCharsets.UTF_8.decode(buffer);
                int i = 0;
                for (; charBuffer.get(i) != '<' || charBuffer.get(i + 1) != 'r'
                    || charBuffer.get(i + 2) != 'o' || charBuffer.get(i + 3) != 'w'
                    || charBuffer.get(i + 4) > ' ' || charBuffer.get(i + 5) != 'r'
                    || charBuffer.get(i + 6) != '=' || charBuffer.get(i + 7) != '"'; i++);
                i += 8;

                int[] info = innerParse(charBuffer, i);
                dimension |= info[0];
                if (info[1] > (rc & 0x7FFF)) {
                    rc |= info[1] & 0x7FFF;
                }
                if (info[2] < (rc >>> 16)) {
                    rc |= info[2] << 16;
                }
            } else dimension |= 1;
        } catch (IOException e) {
            // Ignore error
            logger.debug("", e);
        }
    }

    private int[] innerParse(CharBuffer charBuffer, int i) {
        int row = 0;
        for (; charBuffer.get(i) != '"'; i++)
            row = row * 10 + (charBuffer.get(i) - '0');
        // spans="
        i++;
        for (; charBuffer.get(i) != 's' || charBuffer.get(i + 1) != 'p'
            || charBuffer.get(i + 2) != 'a' || charBuffer.get(i + 3) != 'n'
            || charBuffer.get(i + 4) != 's' || charBuffer.get(i + 5) != '='
            || charBuffer.get(i + 6) != '"'; i++);
        i += 7;
        int cs = 0, ls = 0;
        for (; charBuffer.get(i) != ':'; i++)
            cs = cs * 10 + (charBuffer.get(i) - '0');
        i++;
        for (; charBuffer.get(i) != '"'; i++)
            ls = ls * 10 + (charBuffer.get(i) - '0');
        return new int[] {row, cs, ls};
    }

    /* Found calc */
    private void findCalc(int row, Cell[] cells, int n) {
        if (nonCalc) return;
        int rr = row + 1;
        long r = ((long) rr) << 16;
        int i = Arrays.binarySearch(calc, r);
        if (i < 0) {
            i = ~i;
            if (i >= calc.length) return;
        }
        long a = calc[i];
        if ((int) (a >> 16) != rr) return;

        cells[(((int) a) & 0x7FFF) - 1].f = true;
        int j = 1;
        if (n == -1) n = cells.length;
        n = Math.min(n, calc.length - i);
        for (; j < n; j++) {
            if ((calc[i + j] >> 16) == rr)
                cells[(((int) calc[i + j]) & 0x7FFF) - 1].f = true;
            else break;
        }
    }
}

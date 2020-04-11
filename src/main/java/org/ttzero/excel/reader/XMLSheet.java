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

package org.ttzero.excel.reader;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.ttzero.excel.entity.style.Styles;

import java.io.BufferedReader;
import java.io.IOException;
import java.nio.ByteBuffer;
import java.nio.CharBuffer;
import java.nio.channels.SeekableByteChannel;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardOpenOption;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Iterator;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * The open-xml format Worksheet
 *
 * @author guanquan.wang on 2018-09-22
 */
class XMLSheet implements Sheet {
    final Logger LOGGER = LoggerFactory.getLogger(getClass());

    XMLSheet() { }

    String name;
    int index; // per sheet index of workbook
//        , size = -1; // size of rows per sheet
    Path path;
    /**
     * The Shared String Table
     */
    SharedStrings sst;
    /**
     * The {@link Styles}
     */
    Styles styles;
    int startRow = -1; // row index of data
    HeaderRow header;
    boolean hidden; // state hidden

    // Range address of the used area in the current sheet
    Dimension dimension;


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
     *      -1: unknown size
     * @deprecated use {@link #getDimension()} to getting full range address
     */
    @Deprecated
    @Override
    public int getSize() {
        return dimension != null ? dimension.lastRow - dimension.firstRow + 1 : -1;
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
        return dimension;
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
            Row row = findRow0(this::createHeader);
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
    boolean eof = false, heof = false; // OPTIONS = false
    long mark;

    XMLRow sRow;

    /**
     * Load sheet.xml as BufferedReader
     *
     * @return Sheet
     * @throws IOException if io error occur
     */
    @Override
    public XMLSheet load() throws IOException {
        LOGGER.debug("load {}", path.toString());
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
                    LOGGER.debug("Dimension-Range: {}", l_);
                    Pattern pat = Pattern.compile("([A-Z]+)(\\d+):([A-Z]+)(\\d+)");
                    Matcher mat = pat.matcher(l_);
                    if (mat.matches()) {
                        dimension = Dimension.of(l_);
                        this.startRow = dimension.firstRow;
                    } else {
                        pat = Pattern.compile("([A-Z]+)(\\d+)");
                        mat = pat.matcher(l_);
                        if (mat.matches()) {
                            this.startRow = Integer.parseInt(mat.group(2));
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
            sRow = new XMLRow(sst, styles, this.startRow > 0 ? this.startRow : 1);
        }

        mark = nChar;
        // Deep read if dimension information not write in header
        if (!eof && dimension == null) {
            parseDimension();
        }
        // Create a empty worksheet dimension
        if (dimension == null) {
            dimension = new Dimension(1, (short) 1, 0, (short) 0);
        }

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
                    LOGGER.debug("end of file.");
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

    protected Row findRow0(HeaderRowFunc func) {
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
            LOGGER.error("Read header row error.");
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

        return func.accept(cb, start, nChar - start);
        // row
//        return new XMLRow(sst, styles, this.startRow > 0 ? this.startRow : 1
//            , parseFormula && hasCalc ? this::findCalc : null).with(cb, start, nChar - start);

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
    void parseDimension() {
        try (SeekableByteChannel channel = Files.newByteChannel(path, StandardOpenOption.READ)) {
            long fileSize = Files.size(path);
            int block = (int) Math.min(1 << 11, fileSize);
            ByteBuffer buffer = ByteBuffer.allocate(block);
            byte[] left = null;
            int left_size = 0, i;

            CharBuffer charBuffer;
            boolean eof;
            for (; ;) {
                channel.position(fileSize - block + left_size);
                channel.read(buffer);
                eof = buffer.limit() < block;
                fileSize -= buffer.limit();
                if (left_size > 0) {
                    buffer.limit(block);
                    buffer.put(left, 0, left_size);
                    left_size = 0;
                }
                buffer.flip();

                charBuffer = StandardCharsets.UTF_8.decode(buffer);
                int limit = charBuffer.limit(), c;
                i = limit - 1;

                if(dimension == null) {
                    c = 7;
                    // <row r="
                    for (; i >= c && (charBuffer.get(i) != '"' || charBuffer.get(i - 1) != '='
                        || charBuffer.get(i - 2) != 'r' || charBuffer.get(i - 3) > ' '
                        || charBuffer.get(i - 4) != 'w' || charBuffer.get(i - 5) != 'o'
                        || charBuffer.get(i - 6) != 'r' || charBuffer.get(i - 7) != '<'); i--)
                        ;
                } else {
                    // <mergeCells
                    c = 12;
                    for (; i >= c && (charBuffer.get(i) > ' ' || charBuffer.get(i - 1) != 's'
                        || charBuffer.get(i - 2) != 'l' || charBuffer.get(i - 3) != 'l'
                        || charBuffer.get(i - 4) != 'e' || charBuffer.get(i - 5) != 'C'
                        || charBuffer.get(i - 6) != 'e' || charBuffer.get(i - 7) != 'g'
                        || charBuffer.get(i - 8) != 'r' || charBuffer.get(i - 9) != 'e'
                        || charBuffer.get(i - 10) != 'm' || charBuffer.get(i - 11) != '<'); i--)
                        ;
                }
                // Not Found
                if (i < c) {
                    if (eof || (eof = fileSize <= 0)) break;
                    for (; i < limit && charBuffer.get(i) != '>'; i++) ;
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
                // Found the last row or empty worksheet
                else {
                    break;
                }
                buffer.position(0);
                buffer.limit(buffer.capacity() - left_size);
            }

            if (eof) return;

            charBuffer.position(i);
            // Dimension
            if (dimension == null) {
                parseDim(channel, charBuffer, buffer);
                charBuffer.position(i);
            }

            // Find mergeCells tag
            parseMerge(channel, charBuffer, buffer, block);
        } catch (IOException e) {
            // Ignore error
            LOGGER.debug("", e);
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
            || charBuffer.get(i + 6) != '"'; i++) ;
        i += 7;
        int cs = 0, ls = 0;
        for (; charBuffer.get(i) != ':'; i++)
            cs = cs * 10 + (charBuffer.get(i) - '0');
        i++;
        for (; charBuffer.get(i) != '"'; i++)
            ls = ls * 10 + (charBuffer.get(i) - '0');
        return new int[] {row, cs, ls};
    }

    private void parseDim(SeekableByteChannel channel, CharBuffer charBuffer, ByteBuffer buffer) throws IOException {
        long rr = 0L;
        int rc, i = charBuffer.position();

        int[] info = innerParse(charBuffer, ++i);
        rc = (info[2] << 16) | (info[1] & 0x7FFF);
        rr |= ((long) info[0]) << 32;

        // Skip if the first row is known
        if (mark > 0) {
            channel.position(mark);
            buffer.clear();
            channel.read(buffer);
            buffer.flip();

            charBuffer = StandardCharsets.UTF_8.decode(buffer);
            i = 0;
            for (; charBuffer.get(i) != '<' || charBuffer.get(i + 1) != 'r'
                || charBuffer.get(i + 2) != 'o' || charBuffer.get(i + 3) != 'w'
                || charBuffer.get(i + 4) > ' ' || charBuffer.get(i + 5) != 'r'
                || charBuffer.get(i + 6) != '=' || charBuffer.get(i + 7) != '"'; i++) ;
            i += 8;

            info = innerParse(charBuffer, i);
            rr |= info[0];
            if (info[1] > (rc & 0x7FFF)) {
                rc |= info[1] & 0x7FFF;
            }
            if (info[2] < (rc >>> 16)) {
                rc |= info[2] << 16;
            }
        } else rr |= 1;

        dimension = new Dimension((int) rr, (short) rc, (int) (rr >>> 32), (short) (rc >>> 16));
    }

    void parseMerge(SeekableByteChannel channel, CharBuffer charBuffer, ByteBuffer buffer, int block)
        throws IOException { }

    Row createHeader(char[] cb, int start, int n) {
        return new XMLRow(sst, styles, this.startRow > 0 ? this.startRow : 1).with(cb, start, n);
    }

    @Override
    public XMLCalcSheet asCalcSheet() {
        return !(this instanceof XMLCalcSheet) ? new XMLCalcSheet(this) : (XMLCalcSheet) this;
    }

    @Override
    public XMLMergeSheet asMergeSheet() {
        return !(this instanceof XMLMergeSheet) ? new XMLMergeSheet(this) : (XMLMergeSheet) this;
    }

//    public FullSheet asFullSheet() {
//        throw new IllegalArgumentException();
//    }


    interface HeaderRowFunc {
        Row accept(char[] cb, int start, int n);
    }
}

/**
 * A sub {@link XMLSheet} to parse cell calc
 */
class XMLCalcSheet extends XMLSheet implements CalcSheet {
    private long[] calc; // Array of formula
    XMLCalcSheet() { }
    XMLCalcSheet(XMLSheet sheet) {
        this.name = sheet.name;
        this.index = sheet.index;
        this.path = sheet.path;
        this.sst = sheet.sst;
        this.styles = sheet.styles;
    }

    /**
     * Load sheet.xml as BufferedReader
     *
     * @return Sheet
     * @throws IOException if io error occur
     */
    @Override
    public XMLCalcSheet load() throws IOException {
        super.load();

        if (!eof && !(sRow instanceof XMLCalcRow) && calc != null) {
            sRow = sRow.asCalcRow().setCalcFun(this::findCalc);
        }
        return this;
    }

    /**
     * Setting formula array
     *
     * @param calc array of formula
     */
    XMLCalcSheet setCalc(long[] calc) {
        this.calc = calc;
        return this;
    }

    @Override
    Row createHeader(char[] cb, int start, int n) {
        return new XMLCalcRow(sst, styles, this.startRow > 0 ? this.startRow : 1, this::findCalc).with(cb, start, n);
    }

    /* Found calc */
    private void findCalc(int row, Cell[] cells, int n) {
        long r = ((long) row) << 16;
        int i = Arrays.binarySearch(calc, r);
        if (i < 0) {
            i = ~i;
            if (i >= calc.length) return;
        }
        long a = calc[i];
        if ((int) (a >> 16) != row) return;

        cells[(((int) a) & 0x7FFF) - 1].f = true;
        int j = 1;
        if (n == -1) n = cells.length;
        n = Math.min(n, calc.length - i);
        for (; j < n; j++) {
            if ((calc[i + j] >> 16) == row)
                cells[(((int) calc[i + j]) & 0x7FFF) - 1].f = true;
            else break;
        }
    }

}

/**
 * A sub {@link XMLSheet} to copy value on merge cells
 */
class XMLMergeSheet extends XMLSheet implements MergeSheet {

    // A merge cells grid
    private Grid mergeCells;

    XMLMergeSheet() { }
    XMLMergeSheet(XMLSheet sheet) {
        this.name = sheet.name;
        this.index = sheet.index;
        this.path = sheet.path;
        this.sst = sheet.sst;
        this.styles = sheet.styles;
    }

    /**
     * Load sheet.xml as BufferedReader
     *
     * @return Sheet
     * @throws IOException if io error occur
     */
    @Override
    public XMLMergeSheet load() throws IOException {
        super.load();

        if (mergeCells == null && !eof) {
            parseDimension();
        }

        if (!eof && !(sRow instanceof XMLMergeRow) && mergeCells != null) {
            sRow = sRow.asMergeRow().setCopyValueFunc(this::mergeCell);
        }

        return this;
    }

    /*
    Parse `mergeCells` tag
     */
    @Override
    void parseMerge(SeekableByteChannel channel, CharBuffer charBuffer, ByteBuffer buffer, int block) throws IOException {
        // Find mergeCells tag
        int limit = charBuffer.limit();
        List<Dimension> mergeCells = null;
        boolean eof = false;
        int i = charBuffer.position();
        for (; ;) {
            for (; i < limit - 11 && (charBuffer.get(i) != '<' || charBuffer.get(i + 1) != 'm'
                || charBuffer.get(i + 2) != 'e' || charBuffer.get(i + 3) != 'r'
                || charBuffer.get(i + 4) != 'g' || charBuffer.get(i + 5) != 'e'
                || charBuffer.get(i + 6) != 'C' || charBuffer.get(i + 7) != 'e'
                || charBuffer.get(i + 8) != 'l' || charBuffer.get(i + 9) != 'l'
                || charBuffer.get(i + 10) > ' '); i++) ;

            if (i >= limit - 11) {
                if (eof) break;
                for (; i >= 0 && charBuffer.get(i) != '<'; i--) ;
                if (i > 0) channel.position(channel.position() + i);

                buffer.clear();
                channel.read(buffer);
                buffer.flip();
                if (buffer.limit() <= 0) break;

                charBuffer = StandardCharsets.UTF_8.decode(buffer);
                limit = charBuffer.limit();
                i = 0;
                eof = limit < block;
            } else {
                i += 11;
                for (; i < limit - 5 && (charBuffer.get(i) != 'r' || charBuffer.get(i + 1) != 'e'
                    || charBuffer.get(i + 2) != 'f' || charBuffer.get(i + 3) != '='
                    || charBuffer.get(i + 4) != '"'); i++) ;
                if (i > limit - 5) continue;
                int f = i += 5;
                for (; charBuffer.get(i) != '"'; i++) ;
                if (i > limit) continue;
                char[] chars = new char[i - f];
                charBuffer.position(f);
                charBuffer.get(chars, 0, i - f);
                if (mergeCells == null) mergeCells = new ArrayList<>();
                mergeCells.add(Dimension.of(new String(chars)));
                charBuffer.position(++i);
            }
        }

        if (mergeCells != null) {
            this.mergeCells = GridFactory.create(mergeCells);
            LOGGER.debug("Grid: {}", this.mergeCells);
        }
    }

    private void mergeCell(int row, Cell cell) {
        mergeCells.merge(row, cell);
    }
}
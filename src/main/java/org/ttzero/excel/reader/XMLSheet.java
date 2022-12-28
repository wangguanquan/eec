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
import java.util.stream.Collectors;

/**
 * The open-xml format Worksheet
 *
 * @author guanquan.wang on 2018-09-22
 */
public class XMLSheet implements Sheet {
    final Logger LOGGER = LoggerFactory.getLogger(getClass());

    public XMLSheet() { }

    public XMLSheet(XMLSheet sheet) {
        this.name = sheet.name;
        this.index = sheet.index;
        this.path = sheet.path;
        this.sst = sheet.sst;
        this.styles = sheet.styles;
        this.id = sheet.id;
        this.startRow = sheet.startRow;
        this.header = sheet.header;
        this.hidden = sheet.hidden;
        this.dimension = sheet.dimension;
        this.drawings = sheet.drawings;
        this.reader = sheet.reader;
        this.cb = sheet.cb;
        this.nChar = sheet.nChar;
        this.length = sheet.length;
        this.eof = sheet.eof;
        this.heof = sheet.heof;
        this.mark = sheet.mark;
        this.sRow = sheet.sRow;
        this.lastRowMark = sheet.lastRowMark;
        this.hrf = sheet.hrf;
        this.hrl = sheet.hrl;
    }

    protected String name;
    protected int index; // per sheet index of workbook
//        , size = -1; // size of rows per sheet
    protected Path path;
    protected int id;
    /**
     * The Shared String Table
     */
    protected SharedStrings sst;
    /**
     * The {@link Styles}
     */
    protected Styles styles;
    protected int startRow = -1; // row index of data
    protected HeaderRow header;
    protected boolean hidden; // state hidden

    // Range address of the used area in the current sheet
    protected Dimension dimension;
    // XMLDrawings
    protected Drawings drawings;
    // Header row
    protected int hrf, hrl;


    /**
     * Setting the worksheet name
     *
     * @param name the worksheet name
     */
    protected void setName(String name) {
        this.name = name;
    }

    /**
     * Setting the worksheet xml path
     *
     * @param path the temp path
     */
    protected void setPath(Path path) {
        this.path = path;
    }

    /**
     * Setting the Shared String Table
     *
     * @param sst the {@link SharedStrings}
     */
    protected void setSst(SharedStrings sst) {
        this.sst = sst;
    }

    /**
     * Setting {@link Styles}
     *
     * @param styles the {@link Styles}
     */
    protected void setStyles(Styles styles) {
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

    protected void setIndex(int index) {
        this.index = index;
    }

    /**
     * Returns the worksheet id
     *
     * @return int value
     */
    @Override
    public int getId() {
        return id;
    }

    /**
     * Setting id of worksheet
     *
     * @param id the id of worksheet
     */
    protected void setId(int id) {
        this.id = id;
    }

    /**
     * Setting the {@link Drawings} info
     *
     * @param drawings resource info
     */
    protected void setDrawings(Drawings drawings) {
        this.drawings = drawings;
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
     *
     * @return true: if current sheet is hidden
     */
    @Override
    public boolean isHidden() {
        return hidden;
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
    @Override
    public Sheet header(int fromRowNum, int toRowNum) {
        rangeCheck(fromRowNum, toRowNum);
        this.hrf = fromRowNum;
        this.hrl = toRowNum;
        return this;
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
            Row row = hrf == 0 ? findRow0(this::createHeader) : getHeader(hrf, hrl);
            if (row != null) {
                header = row instanceof HeaderRow ? (HeaderRow) row : row.asHeader();
                sRow.setHr(header);
            }
        } else if (hrl > 0 && hrl > sRow.index) {
            for (Row row = nextRow(); row != null && row.getRowNum() < hrl; row = nextRow()) ;
        }
        return header;
    }

    protected Row getHeader(int fromRowNum, int toRowNum) {
        if (header != null) return header;
        rangeCheck(fromRowNum, toRowNum);
        if (sRow.index > -1 && fromRowNum < sRow.index)
            throw new IndexOutOfBoundsException("Current row num " + sRow.index + " is great than fromRowNum " + fromRowNum + ". Use Sheet#reset() to reset cursor.");
        // Mutable header rows
        if (toRowNum - fromRowNum > 0) {
            Row[] rows = new Row[toRowNum - fromRowNum + 1];
            int i = 0, lc = -1;
            boolean changeLc = false;
            for (Row row = nextRow(); row != null; row = nextRow()) {
                if (row.getRowNum() >= fromRowNum) {
                    if (row.lc > lc) {
                        lc = row.lc;
                        if (i > 0) changeLc = true;
                    }
                    Row r = new Row() { };
                    r.fc = row.fc;
                    r.lc = row.lc;
                    r.index = row.index;
                    r.sst = row.sst;
                    r.cells = row.copyCells();
                    rows[i++] = r;
                }
                if (row.getRowNum() >= toRowNum) break;
            }
            if (changeLc) {
                for (Row row : rows) {
                    row.lc = lc;
                    row.cells = row.copyCells(lc);
                }
            }

            // Parse merged cells
            XMLMergeSheet tmp = new XMLSheet(this).asMergeSheet();
            tmp.reader = null; // Prevent streams from being consumed by mistake
            List<Dimension> mergeCells = tmp.parseMerge();
            if (mergeCells != null) {
                mergeCells = mergeCells.stream().filter(dim -> dim.firstRow < toRowNum || dim.lastRow > fromRowNum).collect(Collectors.toList());
            }

            return new HeaderRow().with(mergeCells, rows);
        }
        // Single row
        else {
            Row row = nextRow();
            for (; row != null && row.getRowNum() < fromRowNum; row = nextRow());
            return new HeaderRow().with(row);
        }
    }

    // Range check
    static void rangeCheck(int fromRowNum, int toRowNum) {
        if (fromRowNum <= 0)
            throw new IndexOutOfBoundsException("fromIndex = " + fromRowNum);
        if (fromRowNum > toRowNum)
            throw new IllegalArgumentException("fromIndex(" + fromRowNum + ") > toIndex(" + toRowNum + ")");
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
    public Sheet bind(Class<?> clazz, Row row) {
        if (row == null) throw new IllegalArgumentException("Specify the bind row must not be null.");
        if (!row.equals(header)) {
            header = row instanceof HeaderRow ? (HeaderRow) row : row.asHeader();
            sRow.setHr(header);
        }
        try {
            header.setClassOnce(clazz);
        } catch (IllegalAccessException | InstantiationException e) {
            throw new ExcelReadException(e);
        }
        return this;
    }

    @Override
    public String toString() {
        return "Sheet id: " + getId() + ", name: " + getName() + ", dimension: " + getDimension();
    }

    /////////////////////////////////Read sheet file/////////////////////////////////
    protected BufferedReader reader;
    protected char[] cb; // buffer
    protected int nChar, length;
    protected boolean eof = false, heof = false; // OPTIONS = false
    protected long mark;

    // Shared row data, Record the current row
    protected XMLRow sRow;
    protected long lastRowMark;

    /**
     * Load sheet.xml as BufferedReader
     *
     * @return Sheet
     * @throws IOException if io error occur
     */
    @Override
    public XMLSheet load() throws IOException {
        // Prevent multiple parsing
        if (sRow != null) {
            reset();
            return this;
        }
        LOGGER.debug("Load {}", path.toString());
        reader = Files.newBufferedReader(path);
        cb = new char[8192];
        nChar = 0;
        int left = 0;
        loopA: for (; ; ) {
            length = reader.read(cb, left, cb.length - left);
            if (length < 11) break;
            // read size
            if (nChar == 0 && startRow < 1) {
                String line = new String(cb, 56, 1024);
                String size = "<dimension ref=\"";
                int index = line.indexOf(size), end = index > 0 ? line.indexOf('"', index += size.length()) : -1;
                if (end > 0) {
                    String l_ = line.substring(index, end);
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

            // Find the last tag '>'
            for (left = nChar; left > 0 && cb[left--] != '>'; );
            if (left > 0) {
                System.arraycopy(cb, left, cb, 0, length - left);
                mark += nChar;
                nChar = 0;
            }
        }
        // Empty sheet
        if (cb[nChar] == '>') {
            mark += length;
            eof = true;
        } else {
            eof = false;
            mark += nChar;
            sRow = createRow().init(sst, styles, this.startRow > 0 ? this.startRow : 1);
        }

        // Deep read if dimension information not write in header
        if (!eof && dimension == null) {
            parseDimension();
        }
        // Create a empty worksheet dimension
        if (dimension == null) {
            dimension = new Dimension(1, (short) 1);
        }
        LOGGER.debug("Dimension-Range: {}", dimension);

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
            if (mark > 0) {
                reader.skip(mark);
                length = reader.read(cb);
            } else {
                loopA:
                for (; ; ) {
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
            }

            boolean eof = length <= 0 || cb[nChar] == '>';
            if (eof) {
                this.heof = true;
                return null;
            }

            int start = nChar;
            A: for (; ;) {
                // Not empty
                for (; nChar < length - 6; nChar++) {
                    if (cb[nChar] == '<' && cb[nChar + 1] == '/' && cb[nChar + 2] == 'r'
                        && cb[nChar + 3] == 'o' && cb[nChar + 4] == 'w' && cb[nChar + 5] == '>') {
                        nChar += 6;
                        break A;
                    }
                }

                /* Load more when not found end of row tag */
                int n;
                char[] _cb = new char[cb.length << 1];
                System.arraycopy(cb, start, _cb, 0, n = length - start);
                cb = _cb;

                try {
                    length = reader.read(cb, n, cb.length - n);
                    // end of file
                    if (length < 0) {
                        reader.close(); // close reader
                        return null;
                    }
                } catch (IOException e) {
                    throw new ExcelReadException("Parse row data error", e);
                }
                start = 0;
                length += n;
            }

            return func.accept(cb, start, nChar - start);
        } catch (IOException e) {
            LOGGER.error("Read header row error.");
            return null;
        }

    }

    /**
     * Iterating each row of data contains header information and blank lines
     *
     * @return a row iterator
     */
    @Override
    public Iterator<Row> iterator() {
        // If the header row number is specified, the header will be parsed first
        if (hrf > 0) getHeader();
        return new RowSetIterator(this::nextRow, false);
    }

    /**
     * Iterating over data rows without header information and blank lines
     *
     * @return a row iterator
     */
    @Override
    public Iterator<Row> dataIterator() {
        // If the header row number is specified, the header will be parsed first
        if (hrf > 0) getHeader();
        // iterator data rows
        Iterator<Row> nIter = new RowSetIterator(this::nextRow, true);
        /*
        If the header is not specified, the first row will be automatically
         used as the header, if there is a header, the row will not be skipped
         */
        if (hrf == 0 && nIter.hasNext()) {
            Row row = nIter.next();
            if (header == null) header = row.asHeader();
            row.setHr(header);
        }
        return nIter;
    }

    /**
     * List all pictures in workbook
     *
     * @return picture list or null if not exists.
     */
    @Override
    public List<Drawings.Picture> listPictures() {
        return drawings != null ? drawings.listPictures(this) : null;
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
        LOGGER.debug("Reset {}", path.toString());
        try {
            hrf = 0;
            hrl = 0;
            header = null;
            sRow.fc = 0;
            sRow.index = sRow.lc = -1;
            // Close the opening reader
            if (reader != null) {
                reader.close();
            }
            if (cb == null) {
                return this.load();
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

    @Override
    public XMLRow createRow() {
        return new XMLRow();
    }

    /*
    If the Dimension information is not write in header,
    Read from tail and look at the line number of the last line
    to confirm the scope of the entire worksheet.
     */
    void parseDimension() {
        try (SeekableByteChannel channel = Files.newByteChannel(path, StandardOpenOption.READ)) {
            long position = Files.size(path);
            final int block = (int) Math.min(1 << 11, position), c = 7;
            ByteBuffer buffer = ByteBuffer.allocate(block);
            byte[] left = null;
            int left_size = 0, i;

            boolean eof, getit = false;
            for (; ;) {
                position -= (block - left_size);
                channel.position(Math.max(0, position));
                channel.read(buffer);
                if (left_size > 0) {
                    buffer.limit(block);
                    buffer.put(left, 0, left_size);
                }
                buffer.flip();
                eof = buffer.limit() < block;

                int limit = buffer.limit();
                i = limit - 1;

                // <row r="
                for (; i >= c && (buffer.get(i) != '"' || buffer.get(i - 1) != '='
                    || buffer.get(i - 2) != 'r' || buffer.get(i - 3) > ' '
                    || buffer.get(i - 4) != 'w' || buffer.get(i - 5) != 'o'
                    || buffer.get(i - 6) != 'r' || buffer.get(i - 7) != '<'); i--)
                    ;

                // Not Found
                if (i < c) {
                    if (eof || (eof = position <= 0)) break;
                    for (; i < limit && buffer.get(i) != '>'; i++) ;
                    i++;
                    if (i < limit - 1) {
                        buffer.position(i);
                        left_size = i;
                        if (left == null || left_size > left.length) {
                            left = new byte[left_size];
                        }
                        buffer.position(0);
                        buffer.get(left, 0, left_size);
                    } else left_size = 0;
                }
                // Found the last row or empty worksheet
                else {
                    getit = true;
                    buffer.position(i);
                    // Align channel and buffer position
                    if (left_size > 0) channel.position(channel.position() + left_size);
                    lastRowMark = channel.position() - buffer.remaining();
                    break;
                }
                buffer.position(0);
                buffer.limit(buffer.capacity() - left_size);
            }

            // Empty worksheet
            if (!getit) {
                dimension = Dimension.of("A1");
                return;
            }

            // Dimension
            parseDim(channel, buffer);
        } catch (IOException e) {
            // Ignore error
            LOGGER.debug("", e);
        }
    }

    private int[] innerParse(ByteBuffer buffer, int i) {
        int row = 0;
        for (; buffer.get(i) != '"'; i++)
            row = row * 10 + (buffer.get(i) - '0');
        // spans="
        i++;
        for (; buffer.limit() - i >= 7 && (buffer.get(i) != 's' || buffer.get(i + 1) != 'p'
            || buffer.get(i + 2) != 'a' || buffer.get(i + 3) != 'n'
            || buffer.get(i + 4) != 's' || buffer.get(i + 5) != '='
            || buffer.get(i + 6) != '"'); i++) ;
        i += 7;
        int cs = 0, ls = 0;
        if (buffer.limit() <= i) {
            eof = true; // EOF
            return new int[] { row, 1, 1 };
        }
        for (; buffer.get(i) != ':'; i++)
            cs = cs * 10 + (buffer.get(i) - '0');
        i++;
        for (; buffer.get(i) != '"'; i++)
            ls = ls * 10 + (buffer.get(i) - '0');
        return new int[] { row, cs, ls };
    }

    private void parseDim(SeekableByteChannel channel, ByteBuffer buffer) throws IOException {
        long rr = 0L;
        int rc, i = buffer.position();

        int[] info = innerParse(buffer, ++i);
        rc = (info[2] << 16) | (info[1] & 0x7FFF);
        rr |= ((long) info[0]) << 32;

        // Skip if the first row is known
        if (mark > 0) {
            channel.position(mark);
            buffer.clear();
            channel.read(buffer);
            buffer.flip();

            i = 0;
            for (; buffer.get(i) != '<' || buffer.get(i + 1) != 'r'
                || buffer.get(i + 2) != 'o' || buffer.get(i + 3) != 'w'
                || buffer.get(i + 4) > ' ' || buffer.get(i + 5) != 'r'
                || buffer.get(i + 6) != '=' || buffer.get(i + 7) != '"'; i++) ;
            i += 8;

            info = innerParse(buffer, i);
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

    Row createHeader(char[] cb, int start, int n) {
        return createRow().init(sst, styles, startRow > 0 ? startRow : 1).with(cb, start, n);
    }

    @Override
    public XMLSheet asSheet() {
        return (this instanceof XMLCalcSheet || this instanceof XMLMergeSheet) ? new XMLSheet(this) : this;
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

    // For debug
    static void watch(ByteBuffer buffer) {
        int p = buffer.position();
        byte[] bytes = new byte[Math.min(1 << 7, buffer.remaining())];
        buffer.get(bytes, 0, bytes.length);
        System.out.println(new String(bytes, 0, bytes.length, StandardCharsets.US_ASCII));
        buffer.position(p);
    }
}

/**
 * A sub {@link XMLSheet} to parse cell calc
 */
class XMLCalcSheet extends XMLSheet implements CalcSheet {
    private long[] calc; // Array of formula
    boolean ready;

    XMLCalcSheet(XMLSheet sheet) {
        this.name = sheet.name;
        this.index = sheet.index;
        this.path = sheet.path;
        this.sst = sheet.sst;
        this.styles = sheet.styles;
        this.id = sheet.id;
        this.startRow = sheet.startRow;
        this.header = sheet.header;
        this.hidden = sheet.hidden;
        this.dimension = sheet.dimension;
        this.drawings = sheet.drawings;
        this.reader = sheet.reader;
        this.cb = sheet.cb;
        this.nChar = sheet.nChar;
        this.length = sheet.length;
        this.eof = sheet.eof;
        this.heof = sheet.heof;
        this.mark = sheet.mark;
        this.sRow = sheet.sRow;
        this.lastRowMark = sheet.lastRowMark;

        if (this.path != null) {

            if (reader != null && !ready) this.load0();
        }
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

        load0();

        return this;
    }

    void load0() {
        if (ready) return;

        // Parse calc.xml
        long[][] calcArray = ExcelReader.parseCalcChain(path.getParent());
        if (calcArray != null && calcArray.length >= id) setCalc(calcArray[id - 1]);

        if (!eof && !(sRow instanceof XMLCalcRow)) {
            sRow = sRow.asCalcRow();
            if (calc != null) ((XMLCalcRow) sRow).setCalcFun(this::findCalc);
        }
        ready = true;
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
    boolean ready;

    XMLMergeSheet(XMLSheet sheet) {
        this.name = sheet.name;
        this.index = sheet.index;
        this.path = sheet.path;
        this.sst = sheet.sst;
        this.styles = sheet.styles;
        this.id = sheet.id;
        this.startRow = sheet.startRow;
        this.header = sheet.header;
        this.hidden = sheet.hidden;
        this.dimension = sheet.dimension;
        this.drawings = sheet.drawings;
        this.reader = sheet.reader;
        this.cb = sheet.cb;
        this.nChar = sheet.nChar;
        this.length = sheet.length;
        this.eof = sheet.eof;
        this.heof = sheet.heof;
        this.mark = sheet.mark;
        this.sRow = sheet.sRow;
        this.lastRowMark = sheet.lastRowMark;

        if (path != null) {
            if (reader != null && !ready) this.load0();
        }
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

        load0();

        return this;
    }

    // Parse merge tag
    void load0() {
        if (ready) return;
        if (mergeCells == null && !eof) {
            List<Dimension> mergeCells = parseMerge();

            if (mergeCells != null && !mergeCells.isEmpty()) {
                this.mergeCells = GridFactory.create(mergeCells);
                LOGGER.debug("Grid: Size: {} ==> {}", this.mergeCells.size(), this.mergeCells);
            }
        }

        if (!eof && !(sRow instanceof XMLMergeRow) && mergeCells != null) {
            sRow = sRow.asMergeRow().setCopyValueFunc(mergeCells, this::mergeCell);
        }
        ready = true;
    }

    /*
    Parse `mergeCells` tag
     */
    List<Dimension> parseMerge() {
        try (SeekableByteChannel channel = Files.newByteChannel(path, StandardOpenOption.READ)) {
            long position = Files.size(path);
            final int block = (int) Math.min(1 << 11, position), c = 12;
            ByteBuffer buffer = ByteBuffer.allocate(block);
            byte[] left;
            int left_size = 0, i, limit;
            boolean eof, getit = false;
            // Skip if marked
            if (lastRowMark > 0L) {
                channel.position(lastRowMark);
                channel.read(buffer);
                buffer.flip();
                getit = true;
            } else {
                left = new byte[12];
                for (; ; ) {
                    channel.position(Math.max(0, position -= block - left_size));
                    channel.read(buffer);
                    if (left_size > 0) {
                        buffer.limit(block);
                        buffer.put(left, 0, left_size);
                    }
                    buffer.flip();
                    eof = buffer.limit() < block;

                    limit = buffer.limit();
                    i = limit - 1;

                    // </sheetData>
                    for (; i >= c && (buffer.get(i) != '>' || buffer.get(i - 1) != 'a'
                        || buffer.get(i - 2) != 't' || buffer.get(i - 3) != 'a'
                        || buffer.get(i - 4) != 'D' || buffer.get(i - 5) != 't'
                        || buffer.get(i - 6) != 'e' || buffer.get(i - 7) != 'e'
                        || buffer.get(i - 8) != 'h' || buffer.get(i - 9) != 's'
                        || buffer.get(i - 10) != '/' || buffer.get(i - 11) != '<'); i--)
                        ;

                    // Not Found
                    if (i < c) {
                        if (eof) break;
                        buffer.position(0);
                        left_size = 12;
                        buffer.get(left, 0, left_size);
                        buffer.clear();
                        buffer.limit(block - left_size);
                    }
                    // Found the last row or empty worksheet
                    else {
                        buffer.position(i + 1);
                        getit = true;
                        // Align channel and buffer position
                        channel.position(channel.position() + left_size);
                        break;
                    }
                }
            }

            if (getit) {
                // Find mergeCells tag
                return parseMerge(channel, buffer, block);
            }
        } catch (IOException e) {
            // Ignore error
            LOGGER.warn("", e);
        }
        return null;
    }

    // Find mergeCell tags
    List<Dimension> parseMerge(SeekableByteChannel channel, ByteBuffer buffer, int block) throws IOException {
        List<Dimension> mergeCells = new ArrayList<>();
        boolean eof = false;
        int limit = buffer.limit(), i = buffer.position(), f, n;
        byte[] bytes = new byte[32];
        for (; ;) {
            for (; i < limit - 11 && (buffer.get(i) != '<' || buffer.get(i + 1) != 'm'
                || buffer.get(i + 2) != 'e' || buffer.get(i + 3) != 'r'
                || buffer.get(i + 4) != 'g' || buffer.get(i + 5) != 'e'
                || buffer.get(i + 6) != 'C' || buffer.get(i + 7) != 'e'
                || buffer.get(i + 8) != 'l' || buffer.get(i + 9) != 'l'
                || buffer.get(i + 10) > ' '); i++) ;

            if (i >= limit - 22) {
                if (eof) break;
                buffer.position(limit - 22);
                for (; i < buffer.limit() && buffer.get() != '<'; i++) ;
                if (buffer.get(buffer.position() - 1) == '<') buffer.position(buffer.position() - 1);
                buffer.compact();
                channel.read(buffer);
                buffer.flip();
                if ((limit = buffer.limit()) <= 0) break;
                i = 0;
                eof = limit < block;
            } else {
                i += 11;
                for (; i < limit - 5 && (buffer.get(i) != 'r' || buffer.get(i + 1) != 'e'
                    || buffer.get(i + 2) != 'f' || buffer.get(i + 3) != '='
                    || buffer.get(i + 4) != '"'); i++) ;
                f = i += 5;
                for (; i < limit && buffer.get(i) != '"'; i++) ;
                if (i >= limit) {
                    buffer.compact();
                    channel.read(buffer);
                    buffer.flip();
                    if ((limit = buffer.limit()) <= 0) break;

                    i = 0;
                    eof = limit < block;
                    continue;
                }
                n = i - f;
                if (n > bytes.length) bytes = new byte[n];

                buffer.position(f);
                buffer.get(bytes, 0, n);

                mergeCells.add(Dimension.of(new String(bytes, 0, n, StandardCharsets.US_ASCII)));
                buffer.get(); // Skip '"'
            }
        }

        return mergeCells;
    }

    private void mergeCell(int row, Cell cell) {
        mergeCells.merge(row, cell);
    }

    @Override
    public Grid getMergeGrid() {
        return mergeCells;
    }

    @Override
    public List<Dimension> getMergeCells() {
        return parseMerge();
    }
}
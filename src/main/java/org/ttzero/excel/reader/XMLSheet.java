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

import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.Reader;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.Iterator;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Collectors;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;

import static org.ttzero.excel.reader.ExcelReader.getEntry;

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
        this.zipFile = sheet.zipFile;
        this.entry = sheet.entry;
        this.option = sheet.option;
    }

    protected String name;
    protected int index; // per sheet index of workbook
//        , size = -1; // size of rows per sheet
    protected String path;
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
    // Data Source
    protected ZipFile zipFile;
    protected ZipEntry entry;
    // Simple properties
    // The low 16 bits are allocated to the header, while the high 16 bits are occupied by the sheet
    protected int option;

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
    protected void setPath(String path) {
        this.path = path;
    }

    /**
     * Setting the source zip file
     *
     * @param zipFile source data
     */
    protected void setZipFile(ZipFile zipFile) {
        this.zipFile = zipFile;
    }

    /**
     * Setting the worksheet zip entry
     *
     * @param entry source data
     */
    protected void setZipEntry(ZipEntry entry) {
        this.entry = entry;
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
        Dimension d = getDimension();
        return d != null ? d.lastRow - d.firstRow + 1 : -1;
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
        return dimension != null ? dimension : (dimension = parseDimension());
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
                header.setOptions(option << 16 >>> 16);
                sRow.setHr(header);
            }
        } else if (hrl > 0 && hrl > sRow.getRowNum()) {
            Row row0 = findRow0(this::createHeader);
            if (row0 != null && row0.getRowNum() < hrl) {
                for (Row row = nextRow(); row != null && row.getRowNum() < hrl; row = nextRow()) ;
            }
        }
        return header;
    }

    protected Row getHeader(int fromRowNum, int toRowNum) {
        if (header != null || eof) return header;
        rangeCheck(fromRowNum, toRowNum);
        if (sRow.getRowNum() > -1 && fromRowNum < sRow.getRowNum())
            throw new IndexOutOfBoundsException("Current row num " + sRow.getRowNum() + " is great than fromRowNum " + fromRowNum + ". Use Sheet#reset() to reset cursor.");
        HeaderRow headerRow;
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
                    r.index = row.getRowNum();
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

            if (i > 0) {
                List<Dimension> mergeCells;
                if (!(this instanceof MergeSheet)) {
                    // Parse merged cells
                    XMLMergeSheet tmp = new XMLSheet(this).asMergeSheet();
                    tmp.reader = null; // Prevent streams from being consumed by mistake
                    mergeCells = tmp.parseMerge();
                } else mergeCells = ((MergeSheet) this).getMergeCells();

                if (mergeCells != null) {
                    mergeCells = mergeCells.stream().filter(dim -> dim.firstRow < toRowNum || dim.lastRow > fromRowNum).collect(Collectors.toList());
                }

                headerRow = new HeaderRow().with(mergeCells, rows).setOptions(option << 16 >>> 16);
            } else headerRow = new HeaderRow().setOptions(option << 16 >>> 16);
        }
        // Single row
        else {
            Row row = nextRow();
            for (; row != null && row.getRowNum() < fromRowNum; row = nextRow());
            headerRow = row != null ? new HeaderRow().with(row).setOptions(option << 16 >>> 16) : new HeaderRow().setOptions(option << 16 >>> 16);
        }
        // Reset metas
        headerRow.styles = styles;
        headerRow.sst = sst;
        return headerRow;
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
            header.setOptions(option << 16 >>> 16);
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
        return "Sheet id: " + getId() + ", name: " + getName();// + ", dimension: " + getDimension();
    }

    /////////////////////////////////Read sheet file/////////////////////////////////
    protected Reader reader;
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
        LOGGER.debug("Load {}", path);
        reader = new InputStreamReader(zipFile.getInputStream(entry), StandardCharsets.UTF_8);
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
            for (int len = length - 12; nChar < len; nChar++) {
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
            if (dimension == null) dimension = new Dimension(1, (short) 1);
        } else {
            eof = false;
            mark += nChar;
            sRow = createRow().init(sst, styles, this.startRow > 0 ? this.startRow : 1);
        }
        if (dimension != null) LOGGER.debug("Dimension-Range: {}", dimension);

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
                    if (dimension == null)
                        dimension = new Dimension(1, (short) Math.max(sRow.fc, 1), Math.max(sRow.getRowNum(), 1), (short) Math.max(sRow.lc, 1));
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
        try (Reader reader = new InputStreamReader(zipFile.getInputStream(entry), StandardCharsets.UTF_8)) {
            if (mark > 0) {
                reader.skip(mark);
                length = reader.read(cb);
            } else {
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

            return func != null ? func.accept(cb, start, nChar - start) : createHeader(cb, start, nChar - start);
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
        return new RowSetIterator(this::nextRow);
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
        Iterator<Row> nIter = new RowSetIterator.NonBlankIterator(this::nextRow);
        /*
        If the header is not specified, the first row will be automatically
         used as the header, if there is a header, the row will not be skipped
         */
        if (hrf == 0 && nIter.hasNext()) {
            Row row = nIter.next();
            if (header == null) header = row.asHeader().setOptions(option << 16 >>> 16);
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
     * Setting header columns preprocessing properties
     *
     * @return this {@link Sheet}
     */
    @Override
    public Sheet setHeaderColumnReadOption(int option) {
        if (option >= 1 << 17)
            LOGGER.warn("Unrecognized options will be discarded");
        int o = option & ((1 << 17) - 1);
        this.option = this.option >>> 16 << 16 | o;
        if (header != null) header.setOptions(o);
        return this;
    }

    /**
     * Returns Header column options
     *
     * @return this {@link Sheet}
     */
    @Override
    public int getHeaderColumnReadOption() {
        return header != null ? header.option : option << 16 >>> 16;
    }

    /**
     * Reset the {@link XMLSheet}'s row index to begging
     *
     * @return the unread {@link XMLSheet}
     */
    @Override
    public XMLSheet reset() {
        LOGGER.debug("Reset {}", path);
        try {
            hrf = 0;
            hrl = 0;
            header = null;
            sRow.fc = 0;
            sRow.index = sRow.lc = -1;
            sRow.from = sRow.to;
            // Close the opening reader
            if (reader != null) {
                reader.close();
            }
            if (cb == null) {
                return this.load();
            }
            // Reload
            reader = new InputStreamReader(zipFile.getInputStream(entry), StandardCharsets.UTF_8);
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
    Dimension parseDimension() {
        try (InputStream is = zipFile.getInputStream(entry)) {
            // Skips specified number of bytes of uncompressed data.
            if (lastRowMark > 0L) is.skip(lastRowMark);

            // Mark
            long mark = 0L, mark0 = 0L;
            int n, offset = 0, limit = 1 << 14, i, len, f, size = 0;
            byte[] buf = new byte[limit], bytes = new byte[10];
            while ((n = is.read(buf, offset, limit - offset)) > 0) {
                mark += n;
                if ((len = n + offset) < 11) {
                    offset = len;
                    continue;
                }
                i = len - 1;
                // Look up from tail
                for (; i > 1 && (buf[i--] != ' ' || buf[i--] != 'c' || buf[i] != '<'); ) ;
                if (i <= 1 && (buf[i] != '<' || buf[i + 1] != 'c')) {
                    offset = 0;
                    continue;
                }
                if (i <= len - 9) {
                    n = i;
                    i += 3;
                    for (; i < len && buf[i] != 'r' && buf[i + 1] != '=' && buf[i + 2] != '"'; i++);
                    if (i < len - 3) {
                        f = i += 3;
                        for (; i < len && buf[i] != '"'; i++);
                        if (i < len) {
                            System.arraycopy(buf, f, bytes, 0, size = i - f);
                            if (buf[len - 1] == '<') {
                                buf[0] = '<';
                                offset = 1;
                            } else offset = 0;
                            mark0 = mark - (len - n);
                            continue;
                        }
                    }
                    i = n;
                }
                if (i < len) System.arraycopy(buf, i, buf, 0, offset = len - i);
            }
            if (lastRowMark < mark0) lastRowMark = mark0;
            if (size > 0) {
                long cr = ExcelReader.cellRangeToLong(new String(bytes, 0, size, StandardCharsets.US_ASCII));
                return new Dimension(1, (short) 1, (int) (cr >> 16), (short) (cr & 0x7FFF));
            }
        } catch (IOException e) {
            // Ignore error
            LOGGER.warn("", e);
        }

        return Dimension.of("A1");
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
        this.zipFile = sheet.zipFile;
        this.entry = sheet.entry;
        this.hrf = sheet.hrf;
        this.hrl = sheet.hrl;
        this.option = sheet.option;

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

        // Parse calcChain.xml
        ZipEntry entry = getEntry(zipFile, "xl/calcChain.xml");
        long[][] calcArray = null;
        try {
            calcArray = entry != null ? ExcelReader.parseCalcChain(zipFile.getInputStream(entry)) : null;
        } catch (IOException e) {
            LOGGER.warn("Parse calcChain failed, formula will be ignored");
        }
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
    List<Dimension> dimensions;

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
        this.zipFile = sheet.zipFile;
        this.entry = sheet.entry;
        this.hrf = sheet.hrf;
        this.hrl = sheet.hrl;
        this.option = sheet.option;

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
                LOGGER.debug("Grid: {} ===> Size: {}", this.mergeCells.getClass(), this.mergeCells.size());
                this.dimensions = mergeCells;
            } else this.dimensions = Collections.emptyList();
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
        List<Dimension> list = null;
        try (InputStream is = zipFile.getInputStream(entry)) {
            // Skips specified number of bytes of uncompressed data.
            if (lastRowMark > 0L) is.skip(lastRowMark);

            int n, offset = 0, limit = 1 << 14, i, len;
            byte[] buf = new byte[limit];
            while ((n = is.read(buf, offset, limit - offset)) > 0) {
                if ((len = n + offset) < 11) {
                    offset = len;
                    continue;
                }
                i = 0; len--;
                for (; i < len && (buf[i++] != '<' || buf[i] != 'm'); ) ;
                // Compact
                if (i >= len) {
                    if (buf[i] == '<') {
                        buf[0] = '<';
                        offset = 1;
                    } else if (buf[i - 1] == '<' && buf[i] == 'm') {
                        buf[0] = '<';
                        buf[1] = 'm';
                        offset = 2;
                    } else offset = 0;
                    continue;
                }
                // Get it
                len++;
                if (len - i < 11) {
                    System.arraycopy(buf, i, buf, 0, offset = len - i);
                    if ((n = is.read(buf, offset, limit - offset)) <= 0)
                        return null;
                    len = n + offset;
                    if (len < 11) {
                        while (((n = is.read(buf, offset, limit - offset)) > 0)) {
                            if ((len = n + offset) < 11) offset = len;
                            else break;
                        }
                    }
                    i = 0;
                }
                if (len < 11 || buf[i] != 'm' || buf[i + 1] != 'e' || buf[i + 2] != 'r'
                    || buf[i + 3] != 'g' || buf[i + 4] != 'e' || buf[i + 5] != 'C' || buf[i + 6] != 'e'
                    || buf[i + 7] != 'l' || buf[i + 8] != 'l' || buf[i + 9] != 's' || buf[i + 10] > ' ')
                    return null;
                i += 11;
                n = len;
                list = new ArrayList<>();
                int f, t;
                do {
                    if (n < 11) {
                        offset += n;
                        continue;
                    }
                    for (; ;) {
                        for (; i < n - 11 && (buf[i] != '<' || buf[i + 1] != 'm'
                            || buf[i + 2] != 'e' || buf[i + 3] != 'r'
                            || buf[i + 4] != 'g' || buf[i + 5] != 'e'
                            || buf[i + 6] != 'C' || buf[i + 7] != 'e'
                            || buf[i + 8] != 'l' || buf[i + 9] != 'l'
                            || buf[i + 10] > ' '); i++) ;

                        if (i >= n - 11) {
                            System.arraycopy(buf, i, buf, 0, offset = n - i);
                            break;
                        }

                        t = i;
                        i += 11;

                        for (; i < n - 5 && (buf[i] != 'r' || buf[i + 1] != 'e' || buf[i + 2] != 'f'
                            || buf[i + 3] != '=' || buf[i + 4] != '"'); i++) ;

                        if (i >= n - 5) {
                            System.arraycopy(buf, t, buf, 0, offset = n - t);
                            break;
                        }

                        f = i += 5;
                        for (; i < n && buf[i] != '"'; i++) ;

                        if (i >= n) {
                            System.arraycopy(buf, t, buf, 0, offset = n - t);
                            break;
                        }

                        list.add(Dimension.of(new String(buf, f, i - f, StandardCharsets.US_ASCII)));
                    }
                    i = 0;
                } while ((n = is.read(buf, offset, limit - offset)) > 0 && (n += offset) > 0);
            }
        } catch (IOException e) {
            // Ignore error
            LOGGER.warn("", e);
        }
        return list;
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
        return dimensions != null ? dimensions : (dimensions = parseMerge());
    }
}
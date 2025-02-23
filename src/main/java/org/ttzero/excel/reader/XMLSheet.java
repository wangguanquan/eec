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

import org.dom4j.Document;
import org.dom4j.DocumentException;
import org.dom4j.DocumentHelper;
import org.dom4j.Element;
import org.dom4j.io.SAXReader;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.ttzero.excel.entity.Comment;
import org.ttzero.excel.entity.Comments;
import org.ttzero.excel.entity.Panes;
import org.ttzero.excel.entity.Relationship;
import org.ttzero.excel.entity.style.Styles;
import org.ttzero.excel.manager.Const;
import org.ttzero.excel.manager.RelManager;
import org.ttzero.excel.util.StringUtil;

import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.Reader;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;

import static org.ttzero.excel.reader.ExcelReader.getEntry;
import static org.ttzero.excel.reader.CalcSheet.parseCalcChain;

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
        this.sRow = (sheet.sRow == null || sheet.sRow.getClass() != XMLRow.class) && !eof ? createRow().init(sst, styles, startRow) : sheet.sRow;
        this.lastRowMark = sheet.lastRowMark;
        this.hrf = sheet.hrf;
        this.hrl = sheet.hrl;
        this.zipFile = sheet.zipFile;
        this.entry = sheet.entry;
        this.option = sheet.option;
        this.relManager = sheet.relManager;
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
    // Relationship Manager
    protected RelManager relManager;

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
    protected void setSharedStrings(SharedStrings sst) {
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
            Row row = hrf == 0 ? findRow0() : getHeader(hrf, hrl);
            if (row != null) {
                header = row instanceof HeaderRow ? (HeaderRow) row : row.asHeader();
                header.setOptions(option << 16 >>> 16);
                sRow.setHeader(header);
            }
        } else if (hrl > 0 && hrl > sRow.getRowNum()) {
            Row row0 = findRow0();
            if (row0 != null && row0.getRowNum() < hrl) {
                for (Row row = nextRow(); row != null && row.getRowNum() < hrl; row = nextRow()) ;
            }
            if (sRow != null && sRow.hr != header) sRow.setHeader(header);
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
                    Map<String, Object> tags = tmp.parseTails();
                    mergeCells = (List<Dimension>) tags.get("mergeCells");
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
            sRow.setHeader(header);
        }
        try {
            header.setClassOnce(clazz);
        } catch (IllegalAccessException | InstantiationException e) {
            throw new ExcelReadException(e);
        }
        return this;
    }

    /**
     * 获取关系管理器
     *
     * @return 关系管理器实例
     */
    public RelManager getRelManager() {
        if (relManager == null) {
            int i = path.lastIndexOf('/');
            if (i < 0) i = path.lastIndexOf('\\');
            String fileName = path.substring(i + 1);
            ZipEntry entry = getEntry(zipFile, "xl/worksheets/_rels/" + fileName + ".rels");
            if (entry != null) {
                SAXReader reader = SAXReader.createDefault();
                try {
                    Document document = reader.read(zipFile.getInputStream(entry));
                    List<Element> list = document.getRootElement().elements();
                    Relationship[] rels = new Relationship[list.size()];
                    i = 0;
                    for (Element e : list) {
                        rels[i++] = new Relationship(e.attributeValue("Id"), e.attributeValue("Target"), e.attributeValue("Type"), e.attributeValue("TargetMode"));
                    }
                    relManager = RelManager.of(rels);
                } catch (DocumentException | IOException e) {
                    LOGGER.error("The file format is incorrect or corrupted. [{}]", entry.getName());
                }
            }
            if (relManager == null) relManager = new RelManager();
        }
        return relManager;
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
     * 加载sheet.xml并解析头信息，如果已加载则直接跳到标记位
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
        nChar = 0; mark = 0;

        // 解析头信息
        parseBOF();

        // Empty sheet
        if (length <= 0) eof = true;
        if (!eof) sRow = createRow().init(sst, styles, this.startRow > 0 ? this.startRow : 1);

        LOGGER.debug("eof: {}, mark: {}", eof, mark);
        if (dimension != null) LOGGER.debug("Dimension-Range: {}", dimension);

        return this;
    }

    // 解析工作表头信息，注意reader的position必须从0开始
    protected void parseBOF() throws IOException {
        int left = 0;
        loopA: while ((length = reader.read(cb, left, cb.length - left)) > 0) {
            if ((length += left) < 11) {
                left = length;
                continue;
            }
            left = 0;

            for (; ;) {
                // 查找起始标签
                for (; nChar < length && cb[nChar] != '<'; nChar++) ;

                if (nChar == length) {
                    mark += nChar;
                    nChar = 0;
                    break;
                }
                int offset = nChar;
                // 跳过结束标签
                if (++nChar < length && cb[nChar] == '/') continue;

                for (; nChar < length && cb[nChar] != '>'; nChar++) ;

                if (nChar == length) {
                    if (offset != 0) {
                        left = nChar - offset;
                        System.arraycopy(cb, offset, cb, 0, left);
                        mark += offset;
                    } else {
                        cb = Arrays.copyOf(cb, cb.length << 1);
                        left = length;
                    }
                    nChar = 0;
                    break;
                }

                int n = ++nChar - offset;
                if (n >= 11 && cb[offset + 1] == 's' && cb[offset + 2] == 'h'
                    && cb[offset + 3] == 'e' && cb[offset + 4] == 'e' && cb[offset + 5] == 't'
                    && cb[offset + 6] == 'D' && cb[offset + 7] == 'a' && cb[offset + 8] == 't'
                    && cb[offset + 9] == 'a' && (cb[offset + 10] == '>' || cb[offset + 10] == '/')) {
                    mark += offset + 11;
                    eof = cb[offset + 10] == '/';
                    break loopA;
                }
                // 如果reader的position不从0开始则遇到<row>就停止否则会一直读取末尾
                if (n >= 5 && cb[offset + 1] == 'r' && cb[offset + 2] == 'o' && cb[offset + 3] == 'w'
                    && (cb[offset + 4] == '>' || cb[offset + 4] == '/')) {
                    mark += offset;
                    eof = false;
                    break loopA;
                }

                // 解析每个子节点
                subElement(cb, offset, n);
            }
        }
    }

    /**
     * iterator rows
     *
     * @return Row
     */
    protected XMLRow nextRow() {
        if (eof) return null;
        boolean endTag = false;
        int start = nChar;
        // find end of row tag
        for (; ++nChar < length && cb[nChar] != '>'; ) ;
        // Empty Row
        if (nChar < length && cb[nChar - 1] == '/') {
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
                if (e.getMessage() != null && e.getMessage().contains("Stream closed")) {
                    eof = true;
                    return null;
                }
                throw new ExcelReadException("Parse row data error", e);
            }
            nChar = 0;
            length += n;
            return nextRow();
        }

        // share row
        return sRow.with(cb, start, nChar - start);
    }

    protected Row findRow0() {
        // 临时保存工作表现有状态
        Marker marker = Marker.of(this);

        Row firstRow = null;
        try {
            load();
            if (!this.eof) {
                XMLRow row = nextRow();
                if (row != null) firstRow = createHeader(row.cb, row.from, row.to - row.from);
            }
            if (this.reader != null) this.reader.close();
        } catch (IOException e) {
            LOGGER.error("Read header row error.");
        }

        this.heof = firstRow == null;

        // 还原工作表状态
        marker.reset();

        return firstRow;
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
            row.setHeader(header);
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
            if (sRow != null) {
                sRow.fc = 0;
                sRow.index = sRow.lc = -1;
                sRow.from = sRow.to;
            }
            // Close the opening reader
            if (reader != null) {
                reader.close();
            }
            if (cb == null) {
                sRow = null; // Repair possible dead cycles
                return this.load();
            }
            // Reload
            reader = new InputStreamReader(zipFile.getInputStream(entry), StandardCharsets.UTF_8);
            reader.skip(mark);
            length = reader.read(cb);
            nChar = 0;
            eof = sRow == null;
        } catch (IOException e) {
            throw new ExcelReadException("Reset worksheet[" + getName() + "] error occur.", e);
        }

        return this;
    }

    public XMLRow createRow() {
        return new XMLRow();
    }

    /*
    If the Dimension information is not write in header,
    Read from tail and look at the line number of the last line
    to confirm the scope of the entire worksheet.
     */
    protected Dimension parseDimension() {
        try (InputStream is = zipFile.getInputStream(entry)) {
            // Skips specified number of bytes of uncompressed data.
            if (lastRowMark > 0L) is.skip(lastRowMark);

            // Mark
            long mark = 0L, mark0 = 0L;
            int n, offset = 0, limit = 1 << 14, i, len, f, row = 1, col = 1;
            byte[] buf = new byte[limit];
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
                            long v = coordinateToLong(buf, f, i - f);
                            row = (int) (v >>> 16);
                            int c = (int) (v & 0x7FFF);
                            // 取列较大的值
                            if (c > col) col = c;
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
            return new Dimension(1, (short) 1, row, (short) col);
        } catch (IOException e) {
            // Ignore error
            LOGGER.warn("", e);
        }

        return Dimension.of("A1");
    }

    // 高48位保存Row，低16位保存Col
    static long coordinateToLong(byte[] buf, int from, int len) {
        long v = 0L;
        int n = 0;
        for (int i = 0; i < len; i++) {
            byte value = buf[i + from];
            if (value >= 'A' && value <= 'Z') {
                v = v * 26 + value - 'A' + 1;
            }
            else if (value >= '0' && value <= '9') {
                n = n * 10 + value - '0';
            }
            else if (value >= 'a' && value <= 'z') {
                v = v * 26 + value - 'a' + 1;
            }
            else break;
        }
        return (v & 0x7FFF) | ((long) n) << 16;
    }

    // 解析感兴趣的子节点
    protected void subElement(char[] cb, int offset, int n) {
        // 这里只处理dimension节点
        if (n < 20) return;
        if (cb[offset + 1] == 'd' && cb[offset + 2] == 'i' && cb[offset + 3] == 'm' && cb[offset + 4] == 'e' && cb[offset + 5] == 'n'
            && cb[offset + 6] == 's' && cb[offset + 7] == 'i' && cb[offset + 8] == 'o' && cb[offset + 9] == 'n') {
            offset += 10;
            int end = offset + n + 1;
            for (; offset < end && cb[offset] != '"'; offset++);
            int i = ++offset;
            for (; offset < end && cb[offset] != '"'; offset++);
            if (offset < end && offset > i) {
                Dimension dim = Dimension.of(new String(cb, i, offset - i));
                if (dim.width > 1 || dim.height > 1) this.dimension = dim;
            }
        }
    }

    protected Row createHeader(char[] cb, int start, int n) {
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

    @Override
    public XMLFullSheet asFullSheet() {
        // XMLCalcSheet 和 XMLMergeSheet 继承至 XMLFullSheet 所以这里没办法使用instanceof判断
        return this.getClass() != XMLFullSheet.class ? new XMLFullSheet(this) : (XMLFullSheet) this;
    }

    /**
     * 保存工作表当前状态并担任reset方法还原
     */
    protected static class Marker {
        private final Reader reader;
        private final char[] cb;
        private final int nChar,length;
        private final boolean eof, heof;
        private final long mark, lastRowMark;
        private final XMLRow sRow;
        private final XMLSheet sheet;

        public Marker(XMLSheet sheet) {
            this.sheet = sheet;
            this.reader = sheet.reader;
            this.cb = sheet.cb;
            this.nChar = sheet.nChar;
            this.length = sheet.length;
            this.eof = sheet.eof;
            this.heof = sheet.heof;
            this.mark = sheet.mark;
            this.lastRowMark = sheet.lastRowMark;
            this.sRow = sheet.sRow;
            sheet.reader = null; // 为了保护工作表的reader不被读取
            sheet.sRow = null;
        }

        public static Marker of(XMLSheet sheet) {
            return new Marker(sheet);
        }

        public void reset() {
            // 关闭流，非必要操作zip流被关闭的时候所有打开的资源都会被一并清除
            if (sheet.reader != null) {
                try {
                    sheet.reader.close();
                } catch (IOException e) { }
            }
            sheet.reader = this.reader;
            sheet.cb = this.cb;
            sheet.nChar = this.nChar;
            sheet.length = this.length;
            sheet.eof = this.eof;
            sheet.heof = this.heof;
            sheet.mark = this.mark;
            sheet.lastRowMark = this.lastRowMark;
            sheet.sRow = this.sRow;
        }
    }
}

/**
 * A sub {@link XMLSheet} to parse cell calc
 */
class XMLCalcSheet extends XMLFullSheet implements CalcSheet {
    XMLCalcSheet(XMLSheet sheet) {
        super(sheet);
    }

    @Override
    void load0() {
        if (ready || eof) return;

        // Parse calcChain.xml
        ZipEntry entry = getEntry(zipFile, "xl/calcChain.xml");
        long[][] calcArray = null;
        try {
            calcArray = entry != null ? parseCalcChain(zipFile.getInputStream(entry)) : null;
        } catch (IOException e) {
            LOGGER.warn("Parse calcChain failed, formula will be ignored");
        }
        if (calcArray != null && calcArray.length >= id) setCalc(calcArray[id - 1]);

        if (!(sRow instanceof XMLCalcRow)) sRow = sRow.asCalcRow();
        if (calc != null) ((XMLCalcRow) sRow).setCalcFun(this::findCalc);
        ready = true;
    }

    @Override
    protected Row createHeader(char[] cb, int start, int n) {
        return createRow().init(sst, styles, this.startRow > 0 ? this.startRow : 1).with(cb, start, n).asCalcRow().setCalcFun(this::findCalc);
    }
}

/**
 * A sub {@link XMLSheet} to copy value on merge cells
 */
class XMLMergeSheet extends XMLFullSheet implements MergeSheet {

    XMLMergeSheet(XMLSheet sheet) {
        super(sheet);
    }

    // Parse merge tag
    @Override
    void load0() {
        if (ready || eof) return;
        if (mergeCells == null) {
            Map<String, Object> tags = parseTails();
            @SuppressWarnings("unchecked")
            List<Dimension> mergeCells = (List<Dimension>) tags.get("mergeCells");

            if (mergeCells != null && !mergeCells.isEmpty()) {
                this.mergeGrid = GridFactory.create(mergeCells);
                this.mergeCells = mergeCells;
                LOGGER.debug("Grid: {} ===> Size: {}", mergeGrid.getClass().getSimpleName(), mergeGrid.size());
            } else {
                this.mergeGrid = new Grid.FastGrid(Dimension.of("A1"));
                this.mergeCells = Collections.emptyList();
            }
        }

        if (!(sRow instanceof XMLMergeRow)) sRow = sRow.asMergeRow();
        ((XMLMergeRow) sRow).setCopyValueFunc(mergeGrid,  mergeGrid::merge);
        ready = true;
    }

    @Override
    protected Row createHeader(char[] cb, int start, int n) {
        return createRow().init(sst, styles, startRow > 0 ? startRow : 1).with(cb, start, n).asMergeRow();
    }
}

/**
 * A sub {@link XMLSheet} to parse all attributes
 */
class XMLFullSheet extends XMLSheet implements FullSheet {
    long[] calc; // Array of formula
    boolean ready, tailPared;
    // A merge cells grid
    Grid mergeGrid;
    List<Dimension> mergeCells;
    int showGridLines = 1; // 默认显示
    Panes panes; // 冻结
    double defaultColWidth = -1D, defaultRowHeight = -1D;
    List<Col> cols; // 列宽
    Dimension filter; // 过滤
    Integer zoomScale; // 缩放比例
    String legacyDrawing;

    XMLFullSheet(XMLSheet sheet) {
        super(sheet);

        if (this.path != null && reader != null && !ready) {
            this.load0();
        }
    }

    /**
     * Load sheet.xml as BufferedReader
     *
     * @return Sheet
     * @throws IOException if io error occur
     */
    @Override
    public XMLFullSheet load() throws IOException {
        super.load();

        load0();

        return this;
    }

    void load0() {
        if (ready || eof) return;

        // 解析公式
        ZipEntry entry = getEntry(zipFile, "xl/calcChain.xml");
        long[][] calcArray = null;
        try {
            calcArray = entry != null ? parseCalcChain(zipFile.getInputStream(entry)) : null;
        } catch (IOException e) {
            LOGGER.warn("Parse calcChain failed, formula will be ignored");
        }
        if (calcArray != null && calcArray.length >= id) setCalc(calcArray[id - 1]);

        if (!(sRow instanceof XMLFullRow)) sRow = sRow.asFullRow();
        if (calc != null) ((XMLFullRow) sRow).setCalcFun(this::findCalc);

        // 默认不复制合并单元格的值
        if (((option >> 17) & 1) == 1 && getMergeGrid() != null) ((XMLFullRow) sRow).setCopyValueFunc(getMergeGrid(), mergeGrid::merge);
        else ((XMLFullRow) sRow).setCopyValueFunc(new Grid.FastGrid(Dimension.of("A1")), (row, cells) -> { });

        ready = true;

        // 再次解析头部（需要解析完整的头部覆写subElement方法
        if (cols == null && defaultRowHeight < 0D && defaultColWidth < 0D && panes == null && showGridLines == 1) {
            Marker marker = Marker.of(this);
            try {
                super.load(); // 这里再次解析不会出现异常
            } catch (IOException e) { }
            marker.reset();
        }
    }

    /**
     * Setting formula array
     *
     * @param calc array of formula
     */
    XMLFullSheet setCalc(long[] calc) {
        this.calc = calc;
        return this;
    }

    @Override
    protected Row createHeader(char[] cb, int start, int n) {
        return ((XMLRow) super.createHeader(cb, start, n)).asFullRow().setCalcFun(this::findCalc);
    }

    /* Found calc */
    void findCalc(int row, Cell[] cells, int n) {
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

    /*
    Parse `mergeCells` tag
    TODO parse dataValidation
     */
    void parseTails() {
        List<Dimension> mergeCells = new ArrayList<>();
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
                i = 0; n = len - 12;
//                for (; i < n && (buf[i] != '<' || ((buf[i + 1] != 'm' || buf[i + 5] != 'e') && (buf[i + 1] != 'a' || buf[i + 5] != 'F') && (buf[i + 1] != 'd' || buf[i + 5] != 'V'))); i++) ;
                for (; i < n && (buf[i] != '<' || buf[i + 1] != '/' || buf[i + 2] != 's' || buf[i + 3] != 'h' || buf[i + 4] != 'e' || buf[i + 5] != 'e' || buf[i + 6] != 't'
                    || buf[i + 7] != 'D' || buf[i + 8] != 'a' || buf[i + 9] != 't' || buf[i + 10] != 'a' || buf[i + 11] != '>'); i++) ;
                // Compact
                if (i >= n) {
                    if (buf[i] == '<') {
                        System.arraycopy(buf, i, buf, 0, offset = len - i);
                    } else offset = 0;
                    continue;
                }
                // Get it
                if (len - i < 11) {
                    System.arraycopy(buf, i, buf, 0, offset = len - i);
                    if ((n = is.read(buf, offset, limit - offset)) <= 0) return;
                    len = n + offset;
                    if (len < 11) {
                        while (((n = is.read(buf, offset, limit - offset)) > 0)) {
                            if ((len = n + offset) < 11) offset = len;
                            else break;
                        }
                    }
                    i = 0;
                }

                if (len < 11) return;
                do {
                    for (; ;) {
                        for (; i < len && buf[i] != '<'; i++) ;
                        if (i == len) {
                            offset = i = 0;
                            break;
                        }
                        int nChar = ++i;
                        for (; nChar < len && buf[nChar] != '>'; nChar++) ;
                        if (nChar == len) {
                            System.arraycopy(buf, i - 1, buf, 0, offset = len - i + 1);
                            i = 0;
                            break;
                        }
                        int length = nChar - i;

                        switch (buf[i]) {
                            // autoFilter
                            case 'a':
                                if (length >= 20 && buf[i + 1] == 'u' && buf[i + 2] == 't' && buf[i + 3] == 'o'
                                    && buf[i + 4] == 'F' && buf[i + 5] == 'i' && buf[i + 6] == 'l' && buf[i + 7] == 't'
                                    && buf[i + 8] == 'e' && buf[i + 9] == 'r' && buf[i + 10] <= ' ') {
                                    i += 11;
                                    for (int k = nChar - 8; i < k && (buf[i] != 'r' || buf[i + 1] != 'e'
                                        || buf[i + 2] != 'f' || buf[i + 3] != '=' || buf[i + 4] != '"'); i++) ;
                                    int a = i += 5;
                                    for (; i < nChar && buf[i] != '"'; i++) ;
                                    if (i > a) filter = Dimension.of(new String(buf, a, i - a, StandardCharsets.US_ASCII));
                                }
                                break;
                            // mergeCells
                            case 'm':
                                if (length >= 20 && buf[i + 1] == 'e' && buf[i + 2] == 'r' && buf[i + 3] == 'g'
                                    && buf[i + 4] == 'e' && buf[i + 5] == 'C' && buf[i + 6] == 'e' && buf[i + 7] == 'l'
                                    && buf[i + 8] == 'l' && buf[i + 9] <= ' ') {
                                    i += 10;
                                    for (int k = nChar - 8; i < k && (buf[i] != 'r' || buf[i + 1] != 'e'
                                        || buf[i + 2] != 'f' || buf[i + 3] != '=' || buf[i + 4] != '"'); i++) ;
                                    int a = i += 5;
                                    for (; i < nChar && buf[i] != '"'; i++) ;
                                    if (i > a) mergeCells.add(Dimension.of(new String(buf, a, i - a, StandardCharsets.US_ASCII)));
                                }
                                break;
                            // dataValidations
                            case 'd':
                                if (len >= 35 && buf[i + 1] == 'a' && buf[i + 2] == 't' && buf[i + 3] == 'a'
                                    && buf[i + 4] == 'V' && buf[i + 5] == 'a' && buf[i + 6] == 'l' && buf[i + 7] == 'i'
                                    && buf[i + 8] == 'd' && buf[i + 9] == 'a' && buf[i + 10] == 't' && buf[i + 11] == 'i'
                                    && buf[i + 12] == 'o' && buf[i + 13] == 'n' && buf[i + 14] <= ' ') {
                                    // TODO
                                }
                                break;
                            // legacyDrawing
                            case 'l':
                                if (len >= 25 && buf[i + 1] == 'e' && buf[i + 2] == 'g' && buf[i + 3] == 'a'
                                    && buf[i + 4] == 'c' && buf[i + 5] == 'y' && buf[i + 6] == 'D' && buf[i + 7] == 'r'
                                    && buf[i + 8] == 'a' && buf[i + 9] == 'w' && buf[i + 10] == 'i' && buf[i + 11] == 'n'
                                    && buf[i + 12] == 'g' && buf[i + 13] <= ' ') {
                                    i += 13;
                                    for (int k = nChar - 8; i < k && buf[i] != 'r' && buf[i + 1] != ':'
                                        && buf[i + 2] != 'i' && buf[i + 3] != 'd' && buf[i + 4] != '=' && buf[i + 5] != '"'; i++) ;
                                    int a = i += 6;
                                    for (; i < nChar && buf[i] != '"'; i++) ;
                                    if (i > a) legacyDrawing = new String(buf, a, i - a, StandardCharsets.US_ASCII);
                                }
                                break;
                        }
                    }
                } while ((len = is.read(buf, offset, limit - offset)) > 0 && (len += offset) > 0);
            }
        } catch (IOException e) {
            // Ignore error
            LOGGER.warn("", e);
        }
        this.mergeCells = mergeCells;
        tailPared = true;
    }

    @Override
    public Grid getMergeGrid() {
        if (mergeGrid != null) return mergeGrid;
        List<Dimension> dims = getMergeCells();
        if (dims != null) {
            mergeGrid = GridFactory.create(dims);
            LOGGER.debug("Grid: {} ===> Size: {}", mergeGrid.getClass().getSimpleName(), mergeGrid.size());
        }
        return mergeGrid;
    }

    @Override
    public List<Dimension> getMergeCells() {
        if (mergeCells == null) parseTails();
        return mergeCells != null && !mergeCells.isEmpty() ? mergeCells : null;
    }

    @Override
    protected void subElement(char[] cb, int offset, int n) {
        String v = new String(cb, offset, n);
        // 去掉不必要的命名空间
        v = v.replace("x14ac:", "").replace("r:", "").replace("mc:", "");
        if (cb[offset + n - 2] == '/') {
            try {
                Document doc = DocumentHelper.parseText(v);
                Element e = doc.getRootElement();
                switch (e.getName()) {
                    case "dimension":
                        String ref = e.attributeValue("ref");
                        Dimension dim;
                        if (StringUtil.isNotEmpty(ref) && ((dim = Dimension.of(ref)).width > 1 || dim.height > 1)) {
                            dimension = dim;
                        }
                        break;
                    case "col":
                        String min = e.attributeValue("min"), max = e.attributeValue("max"), width = e.attributeValue("width")
                            , hidden = e.attributeValue("hidden"), style = e.attributeValue("style");
                        if (cols == null) cols = new ArrayList<>();
                        Col col = new Col(Integer.parseInt(min), Integer.parseInt(max), Double.parseDouble(width), "1".equals(hidden));
                        if (StringUtil.isNotEmpty(style)) col.styleIndex = toInt(style.toCharArray(), 0, style.length());
                        cols.add(col);
                        break;
                    case "pane":
                        String xSplit = e.attributeValue("xSplit"), ySplit = e.attributeValue("ySplit");
                        if (StringUtil.isNotEmpty(ySplit) && Row.testNumberType(ySplit.toCharArray(), 0, ySplit.length()) == 1) panes = Panes.row(Integer.parseInt(ySplit));
                        if (StringUtil.isNotEmpty(xSplit) && Row.testNumberType(xSplit.toCharArray(), 0, xSplit.length()) == 1) {
                            int colIdx = Integer.parseInt(xSplit);
                            if (panes != null) panes.col = colIdx;
                            else panes = Panes.col(colIdx);
                        }
                        break;
                    case "sheetFormatPr":
                        String defaultColWidth = e.attributeValue("defaultColWidth"), defaultRowHeight = e.attributeValue("defaultRowHeight");
                        if (StringUtil.isNotEmpty(defaultColWidth) && Row.testNumberType(defaultColWidth.toCharArray(), 0, defaultColWidth.length()) > 0)
                            this.defaultColWidth = Double.parseDouble(defaultColWidth);
                        if (StringUtil.isNotEmpty(defaultRowHeight) && Row.testNumberType(defaultRowHeight.toCharArray(), 0, defaultRowHeight.length()) > 0)
                            this.defaultRowHeight = Double.parseDouble(defaultRowHeight);
                        break;
                    case "sheetView":
                        String showGridLines = e.attributeValue("showGridLines"), zoomScale = e.attributeValue("zoomScale");
                        if ("0".equals(showGridLines)) this.showGridLines = 0;
                        if (StringUtil.isNotEmpty(zoomScale)) {
                            try {
                                this.zoomScale = Integer.parseInt(zoomScale.trim());
                            } catch (NumberFormatException ex) { }
                        }
                        break;
                }
            } catch (DocumentException e) {
                LOGGER.warn("Parse header tag [" + v + "] failed.", e);
            }
        } else if (v.startsWith("<sheetView") && v.charAt(10) <= ' ') {
            char[] ncb = new char[n + 1];
            System.arraycopy(cb, offset, ncb, 0, n);
            ncb[n - 1] = '/'; ncb[n] = '>';
            try {
                Document doc = DocumentHelper.parseText(new String(ncb, 0, n + 1));
                Element e = doc.getRootElement();
                String showGridLines = e.attributeValue("showGridLines");
                if ("0".equals(showGridLines)) this.showGridLines = 0;
            } catch (DocumentException e) {
                LOGGER.warn("Parse header tag [" + v + "] failed.", e);
            }
        }
    }

    @Override
    public FullSheet copyOnMerged() {
        if (sRow != null && getMergeGrid() != null) ((XMLFullRow) sRow).setCopyValueFunc(getMergeGrid(), mergeGrid::merge);
        else option |= 1 << 17;
        return this;
    }

    @Override
    public Panes getFreezePanes() {
        return panes;
    }

    @Override
    public List<Col> getCols() {
        return cols;
    }

    @Override
    public Dimension getFilter() {
        // 如果filter为则且未解析则重新解析
        if (filter == null && !tailPared) parseTails();
        return filter;
    }

    @Override
    public boolean isShowGridLines() {
        return showGridLines == 1;
    }

    @Override
    public double getDefaultColWidth() {
        return defaultColWidth;
    }

    @Override
    public double getDefaultRowHeight() {
        return defaultRowHeight;
    }

    @Override
    public Integer getZoomScale() {
        return zoomScale;
    }

    @Override
    public Map<Long, Comment> getComments() {
        if (comments == null) {
            RelManager relManager = getRelManager();
            Relationship commentsRel = relManager != null ? relManager.getByType(Const.Relationship.COMMENTS) : null;
            if (commentsRel != null) {
                if (mergeCells == null) getMergeCells();
                Relationship vmlRel = StringUtil.isNotEmpty(legacyDrawing) ? relManager.getById(legacyDrawing) : null;
                if (vmlRel != null) {
                    ZipEntry commentEntry = getEntry(zipFile, "xl/" + toZipPath(commentsRel.getTarget())), vmlEntry = getEntry(zipFile, "xl/" + toZipPath(vmlRel.getTarget()));
                    if (commentEntry != null) {
                        try {
                            comments = Comments.parseComments(zipFile.getInputStream(commentEntry), vmlEntry != null ? zipFile.getInputStream(vmlEntry): null);
                        } catch (IOException ex) {
                            throw new ExcelReadException(ex);
                        }
                    }
                }
            }
            if (comments == null) comments = Collections.emptyMap();
        }
        return comments;
    }
}
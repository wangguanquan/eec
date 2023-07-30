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

package org.ttzero.excel.entity.e7;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.ttzero.excel.entity.IDrawingsWriter;
import org.ttzero.excel.entity.Picture;
import org.ttzero.excel.manager.RelManager;
import org.ttzero.excel.manager.TopNS;
import org.ttzero.excel.entity.Column;
import org.ttzero.excel.entity.Comments;
import org.ttzero.excel.entity.ExcelWriteException;
import org.ttzero.excel.entity.IWorksheetWriter;
import org.ttzero.excel.entity.Panes;
import org.ttzero.excel.entity.Relationship;
import org.ttzero.excel.entity.Row;
import org.ttzero.excel.entity.RowBlock;
import org.ttzero.excel.entity.SharedStrings;
import org.ttzero.excel.entity.Sheet;
import org.ttzero.excel.manager.Const;
import org.ttzero.excel.reader.Cell;
import org.ttzero.excel.reader.Dimension;
import org.ttzero.excel.reader.Grid;
import org.ttzero.excel.reader.GridFactory;
import org.ttzero.excel.util.ExtBufferedWriter;
import org.ttzero.excel.util.FileSignatures;
import org.ttzero.excel.util.FileUtil;
import org.ttzero.excel.util.StringUtil;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.io.ObjectInputStream;
import java.io.ObjectOutputStream;
import java.io.OutputStream;
import java.io.OutputStreamWriter;
import java.lang.reflect.InvocationTargetException;
import java.math.BigDecimal;
import java.net.HttpURLConnection;
import java.net.URL;
import java.nio.ByteBuffer;
import java.nio.ByteOrder;
import java.nio.channels.SeekableByteChannel;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.nio.file.StandardOpenOption;
import java.util.List;
import java.util.function.Supplier;

import static org.ttzero.excel.entity.Sheet.int2Col;
import static org.ttzero.excel.reader.Cell.BINARY;
import static org.ttzero.excel.reader.Cell.BLANK;
import static org.ttzero.excel.reader.Cell.BOOL;
import static org.ttzero.excel.reader.Cell.CHARACTER;
import static org.ttzero.excel.reader.Cell.DATE;
import static org.ttzero.excel.reader.Cell.DATETIME;
import static org.ttzero.excel.reader.Cell.DECIMAL;
import static org.ttzero.excel.reader.Cell.DOUBLE;
import static org.ttzero.excel.reader.Cell.FILE;
import static org.ttzero.excel.reader.Cell.INLINESTR;
import static org.ttzero.excel.reader.Cell.INPUT_STREAM;
import static org.ttzero.excel.reader.Cell.LONG;
import static org.ttzero.excel.reader.Cell.NUMERIC;
import static org.ttzero.excel.reader.Cell.REMOTE_URL;
import static org.ttzero.excel.reader.Cell.SST;
import static org.ttzero.excel.reader.Cell.TIME;
import static org.ttzero.excel.util.ExtBufferedWriter.stringSize;
import static org.ttzero.excel.entity.IWorksheetWriter.isBigDecimal;
import static org.ttzero.excel.entity.IWorksheetWriter.isBool;
import static org.ttzero.excel.entity.IWorksheetWriter.isString;
import static org.ttzero.excel.entity.IWorksheetWriter.isChar;
import static org.ttzero.excel.entity.IWorksheetWriter.isInt;
import static org.ttzero.excel.entity.IWorksheetWriter.isDate;
import static org.ttzero.excel.entity.IWorksheetWriter.isDateTime;
import static org.ttzero.excel.entity.IWorksheetWriter.isTime;
import static org.ttzero.excel.entity.IWorksheetWriter.isFloat;
import static org.ttzero.excel.entity.IWorksheetWriter.isDouble;
import static org.ttzero.excel.entity.IWorksheetWriter.isLong;
import static org.ttzero.excel.entity.IWorksheetWriter.isLocalDate;
import static org.ttzero.excel.entity.IWorksheetWriter.isLocalDateTime;
import static org.ttzero.excel.entity.IWorksheetWriter.isLocalTime;
import static org.ttzero.excel.util.FileUtil.exists;
import static org.ttzero.excel.util.StringUtil.isNotEmpty;

/**
 * @author guanquan.wang at 2019-04-22 16:31
 */
@TopNS(prefix = {"", "r"}, value = "worksheet"
    , uri = {Const.SCHEMA_MAIN, Const.Relationship.RELATIONSHIP})
public class XMLWorksheetWriter implements IWorksheetWriter {
    /**
     * LOGGER
     */
    protected final Logger LOGGER = LoggerFactory.getLogger(getClass());

    // the storage path
    protected Path workSheetPath, mediaPath;
    protected ExtBufferedWriter bw;
    protected Sheet sheet;
    protected Column[] columns;
    protected SharedStrings sst;
    protected Comments comments;
    protected int startRow // The first data-row index
        , startHeaderRow // The first header row index
        , totalRows
        , sheetDataReady // Temporary code to increase compatibility with older versions
        ;
    /**
     * If there are any auto-width columns
     */
    protected boolean includeAutoWidth;
    /**
     * Picture&Chart Support
     */
    protected IDrawingsWriter drawingsWriter;

    public XMLWorksheetWriter() { }

    public XMLWorksheetWriter(Sheet sheet) {
        this.sheet = sheet;
        this.sst = sheet.getSst();
    }

    /**
     * Write a row block
     *
     * @param supplier a row-block supplier
     * @throws IOException if I/O error occur
     */
    @Override
    public void writeTo(Path path, Supplier<RowBlock> supplier) throws IOException {
        Path sheetPath = initWriter(path);

        // Get the first block
        RowBlock rowBlock = supplier.get();

        // write before
        writeBefore();

        // Write body data
        beforeSheetData(sheet.getNonHeader() == 1);

        if (rowBlock != null && rowBlock.hasNext()) {
            do {
                // write row-block data
                writeRowBlock(rowBlock);
                // end of row
                if (rowBlock.isEOF()) break;
            } while ((rowBlock = supplier.get()) != null);
        }

        totalRows = rowBlock != null ? rowBlock.getTotal() : 0;

        // write end
        writeAfter(totalRows);

        // Write some final info
        sheet.afterSheetAccess(workSheetPath);

        // Resize if include auto-width column
        if (includeAutoWidth) {
            // close writer before resize
            close();
            resizeColumnWidth(sheetPath.toFile(), totalRows);
        }
    }

    /**
     * Write a row block
     *
     * @param path the storage path
     * @throws IOException if I/O error occur
     */
    @Override
    public void writeTo(Path path) throws IOException {
        Path sheetPath = initWriter(path);

        // Get the first block
        RowBlock rowBlock = sheet.nextBlock();

        // write before
        writeBefore();

        // Write body data
        beforeSheetData(sheet.getNonHeader() == 1);

        if (rowBlock.hasNext()) {
            for (; ; ) {
                // write row-block data
                writeRowBlock(rowBlock);
                // end of row
                if (rowBlock.isEOF()) break;
                // Get the next block
                rowBlock = sheet.nextBlock();
            }
        }

        totalRows = rowBlock.getTotal();

        // write end
        writeAfter(totalRows);

        // Write some final info
        sheet.afterSheetAccess(workSheetPath);

        // Resize if include auto-width column
        if (includeAutoWidth) {
            // Close writer before resize
            close();
            resizeColumnWidth(sheetPath.toFile(), totalRows);
        }
    }

    protected Path initWriter(Path root) throws IOException {
        this.workSheetPath = root.resolve("worksheets");
        if (!exists(this.workSheetPath)) {
            FileUtil.mkdir(workSheetPath);
        }
        sheet.what("0010", sheet.getName());

        Path sheetPath = workSheetPath.resolve(sheet.getFileName());

        this.bw = new ExtBufferedWriter(Files.newBufferedWriter(
            sheetPath, StandardCharsets.UTF_8));

        if (sst == null) this.sst = sheet.getSst();

        // Check the first row index
        startHeaderRow = sheet.getStartRowIndex();
        if (startHeaderRow <= 0)
            throw new IndexOutOfBoundsException("The start row index must be greater than 0, current = " + startHeaderRow);
        if (getRowLimit() <= startHeaderRow)
            throw new IndexOutOfBoundsException("The start row index must be less than row-limit, current(" + startHeaderRow + ") >= limit(" + getRowLimit() + ")");
        startRow = startHeaderRow;
        return sheetPath;
    }

    /**
     * Rest worksheet
     *
     * @param sheet the worksheet
     * @return self
     */
    @Override
    public IWorksheetWriter setWorksheet(Sheet sheet) {
        this.sheet = sheet;
        return this;
    }

    @Override
    public IWorksheetWriter clone() {
        IWorksheetWriter copy;
        try {
            copy = (IWorksheetWriter) super.clone();
        } catch (CloneNotSupportedException e) {
            ObjectOutputStream oos = null;
            ObjectInputStream ois = null;
            try {
                ByteArrayOutputStream bos = new ByteArrayOutputStream();
                oos = new ObjectOutputStream(bos);
                oos.writeObject(this);

                ois = new ObjectInputStream(new ByteArrayInputStream(bos.toByteArray()));
                copy = (IWorksheetWriter) ois.readObject();
            } catch (IOException | ClassNotFoundException e1) {
                try {
                    copy = getClass().getConstructor(Sheet.class).newInstance(sheet);
                } catch (NoSuchMethodException | IllegalAccessException
                    | InstantiationException | InvocationTargetException e2) {
                    throw new ExcelWriteException(e2);
                }
            } finally {
                FileUtil.close(oos);
                FileUtil.close(ois);
            }
        }
        return copy;
    }

    /**
     * Override this method to modify the maximum number
     * of rows per page, this value contains the header
     * row and data rows
     * <p>
     * eq: limit is 100 means data has 99 rows
     *
     * @return the row limit
     */
    @Override
    public int getRowLimit() {
        return Const.Limit.MAX_ROWS_ON_SHEET;
    }

    /**
     * The Worksheet column limit
     *
     * @return the limit
     */
    @Override
    public int getColumnLimit() {
        return Const.Limit.MAX_COLUMNS_ON_SHEET;
    }

    /**
     * Write worksheet header data
     *
     * @throws IOException if I/O error occur
     */
    protected void writeBefore() throws IOException {
        // The header columns
        columns = sheet.getAndSortHeaderColumns();
        boolean nonHeader = sheet.getNonHeader() == 1;

        bw.write(Const.EXCEL_XML_DECLARATION);
        // Declaration
        bw.newLine();
        // Root node
        writeRootNode();

        // Dimension
        writeDimension();

        // SheetViews default value
        writeSheetViews();

        // Default row height and width
        int fillSpace = 17; // Column width xxx.xx (6byte) + hidden property (11byte)
        BigDecimal width = BigDecimal.valueOf(!nonHeader ? sheet.getDefaultWidth() : 8D).add(new BigDecimal("0.65"));
        // Overflow column width limit
        if (width.compareTo(new BigDecimal(Const.Limit.COLUMN_WIDTH)) > 0) {
            width = new BigDecimal(Const.Limit.COLUMN_WIDTH);
        }
        String defaultWidth = width.setScale(2, BigDecimal.ROUND_HALF_UP).toString();
//        writeSheetFormat(fillSpace, defaultWidth);

        // cols
        writeCols(fillSpace, defaultWidth);

        // Initialization DrawingsWriter
        boolean hasMedia = false;
        for (Column column : columns) {
            hasMedia = column.getColumnType() == 1;
            if (hasMedia) break;
        }
        if (hasMedia) {
            if (mediaPath == null) mediaPath = Files.createDirectories(workSheetPath.getParent().resolve("media"));
            if (drawingsWriter == null) {
                drawingsWriter = createDrawingsWriter();
            }
        }
    }

    /**
     * Write the header row
     *
     * @return row number
     * @throws IOException if I/O error occur
     */
    protected int writeHeaderRow() throws IOException {
        // Write header
        int rowIndex = 0, subColumnSize = columns[0].subColumnSize(), defaultStyleIndex = sheet.defaultHeadStyleIndex();
        Column[][] columnsArray = new Column[columns.length][];
        for (int i = 0; i < columns.length; i++) {
            columnsArray[i] = columns[i].toArray();

            // Free memory after write
//            columns[i].trimTail();
        }
        // Merge cells if exists
        @SuppressWarnings("unchecked")
        List<Dimension> mergeCells = (List<Dimension>) sheet.getExtPropValue(Const.ExtendPropertyKey.MERGE_CELLS);
        Grid mergedGrid = mergeCells != null && !mergeCells.isEmpty() ? GridFactory.create(mergeCells) : null;
        for (int i = subColumnSize - 1; i >= 0; i--) {
            // Custom row height
            double ht = getHeaderHeight(columnsArray, i);
            if (ht < 0) ht = sheet.getHeaderRowHeight();
            int row = startRow(rowIndex++, columns.length, ht);

            String name;
            for (int j = 0, c = 0; j < columns.length; j++) {
                Column hc = columnsArray[j][i];
                name = isNotEmpty(hc.getName()) ? hc.getName() : mergedGrid != null && mergedGrid.test(i + 1, hc.getRealColIndex()) && !isFirstMergedCell(mergeCells, i + 1, hc.getRealColIndex()) ? null : hc.key;
                writeString(name, row, c++, hc.getHeaderStyleIndex() == -1 ? defaultStyleIndex : hc.getHeaderStyleIndex());
            }

            // Write header comments
            for (int j = 0; j < columns.length; j++) {
                Column hc = columnsArray[j][i];
                if (hc.headerComment != null) {
                    if (comments == null) comments = sheet.createComments();
                    comments.addComment(new String(int2Col(hc.getRealColIndex())) + row, hc.headerComment);
                }
            }
            bw.write("</row>");
        }
        return subColumnSize;
    }

    /**
     * Write at after worksheet body
     *
     * @param total the total of rows
     * @throws IOException if I/O error occur
     */
    protected void writeAfter(int total) throws IOException {
        // End target --sheetData
        bw.write("</sheetData>");

        // Merge cells
        writeMergeCells();

        // Others
        afterSheetData();

        // End target
        if (getClass().isAnnotationPresent(TopNS.class)) {
            TopNS topNS = getClass().getAnnotation(TopNS.class);
            bw.write("</");
            bw.write(topNS.value());
            bw.write('>');
        } else {
            bw.write("</worksheet>");
        }
        sheet.what("0009", sheet.getName(), String.valueOf(total));
    }

    /**
     * Write a row-block
     *
     * @param rowBlock the row-block
     * @throws IOException if I/O error occur.
     */
    protected void writeRowBlock(RowBlock rowBlock) throws IOException {
        for (; rowBlock.hasNext(); writeRow(rowBlock.next())) ;
    }

    /**
     * Write a row-block as auto size
     *
     * @param rowBlock the row-block
     * @throws IOException if I/O error occur.
     */
    @Deprecated
    private void writeAutoSizeRowBlock(RowBlock rowBlock) throws IOException {
        for (; rowBlock.hasNext(); writeRowAutoSize(rowBlock.next())) ;
    }

    /**
     * Write begin of row
     *
     * @param rows    the row index (zero base)
     * @param columns the column length
     * @return the row index (one base)
     * @throws IOException if I/O error occur
     * @deprecated replace with {@link #startRow(int, int, double)}
     */
    @Deprecated
    protected int startRow(int rows, int columns) throws IOException {
        // Row number
        int r = rows + startRow;
        // logging
        if (r % 1_0000 == 0) {
            sheet.what("0014", String.valueOf(r));
        }

        bw.write("<row r=\"");
        bw.writeInt(r);
        // default data row height 16.5
        bw.write("\" spans=\"1:");
        bw.writeInt(columns);
        bw.write("\">");
        return r;
    }

    /**
     * Write begin of row
     *
     * @param rows    the row index (zero base)
     * @param columns the column length
     * @param rowHeight the row height
     * @return the row index (one base)
     * @throws IOException if I/O error occur
     */
    protected int startRow(int rows, int columns, double rowHeight) throws IOException {
        // Row number
        int r = rows + startRow;
        // logging
        if (r % 1_0000 == 0) {
            sheet.what("0014", String.valueOf(r));
        }

        bw.write("<row r=\"");
        bw.writeInt(r);
        // default data row height 16.5
        if (rowHeight >= 0D) {
            bw.write("\" customHeight=\"1\" ht=\"");
            bw.write(rowHeight);
        }
        bw.write("\" spans=\"");
        bw.writeInt(this.columns[0].realColIndex);
        bw.write(':');
        bw.writeInt(this.columns[this.columns.length - 1].realColIndex);
        bw.write("\">");
        return r;
    }

    /**
     * Write a row data
     *
     * @param row a row data
     * @throws IOException if I/O error occur
     */
    protected void writeRow(Row row) throws IOException {
        Cell[] cells = row.getCells();
        int len = cells.length, r = startRow(row.getIndex(), len, row.getHeight());

        for (int i = 0; i < len; i++) {
            Cell cell = cells[i];
            int xf = cell.xf;
            switch (cell.t) {
                case INLINESTR:
                case SST:          writeString(cell.sv, r, i, xf);     break;
                case NUMERIC:      writeNumeric(cell.nv, r, i, xf);    break;
                case LONG:         writeNumeric(cell.lv, r, i, xf);    break;
                case DATE:
                case DATETIME:
                case DOUBLE:
                case TIME:         writeDouble(cell.dv, r, i, xf);     break;
                case BOOL:         writeBool(cell.bv, r, i, xf);       break;
                case DECIMAL:      writeDecimal(cell.mv, r, i, xf);    break;
                case CHARACTER:    writeChar(cell.cv, r, i, xf);       break;
                case REMOTE_URL:   writeRemoteMedia(cell.sv, r, i, xf);break;
                case BINARY:       writeBinary(cell.binary, r, i, xf); break;
                case FILE:         writeFile(cell.path, r, i, xf);     break;
                case INPUT_STREAM: writeStream(cell.isv, r, i, xf);    break;
                case BLANK:        writeNull(r, i, xf);                break;
                default:
            }
        }
        bw.write("</row>");
    }


    /**
     * Write a row data
     *
     * @param row a row data
     * @throws IOException if I/O error occur
     */
    @Deprecated
    protected void writeRowAutoSize(Row row) throws IOException {
        Cell[] cells = row.getCells();
        int len = cells.length, r = startRow(row.getIndex(), len, row.getHeight());

        for (int i = 0; i < len; i++) {
            Cell cell = cells[i];
            int xf = cell.xf;
            switch (cell.t) {
                case INLINESTR:
                case SST:       writeStringAutoSize(cell.sv, r, i, xf);  break;
                case NUMERIC:   writeNumericAutoSize(cell.nv, r, i, xf); break;
                case LONG:      writeNumericAutoSize(cell.lv, r, i, xf); break;
                case DATE:
                case DATETIME:
                case DOUBLE:
                case TIME:      writeDoubleAutoSize(cell.dv, r, i, xf);  break;
                case BOOL:      writeBool(cell.bv, r, i, xf);            break;
                case DECIMAL:   writeDecimalAutoSize(cell.mv, r, i, xf); break;
                case CHARACTER: writeChar(cell.cv, r, i, xf);            break;
                case BLANK:     writeNull(r, i, xf);                     break;
                default:
            }
        }
        bw.write("</row>");
    }

    /**
     * Write string value
     *
     * @param s      the string value
     * @param row    the row index
     * @param column the column index
     * @param xf     the style index
     * @throws IOException if I/O error occur
     */
    protected void writeString(String s, int row, int column, int xf) throws IOException {
        // The limit characters per cell check
        if (s != null && s.length() > Const.Limit.MAX_CHARACTERS_PER_CELL) {
            throw new ExcelWriteException("Characters per cell out of limit. size=" + s.length()
                + ", limit=" + Const.Limit.MAX_CHARACTERS_PER_CELL);
        }
        Column hc = columns[column];
        bw.write("<c r=\"");
        bw.write(int2Col(hc.getRealColIndex()));
        bw.writeInt(row);
        int i;
        if (StringUtil.isEmpty(s)) {
            bw.write("\" s=\"");
            bw.writeInt(xf);
            bw.write("\"/>");
        }
        else if (hc.isShare() && (i = sst.get(s)) >= 0) {
            bw.write("\" t=\"s\" s=\"");
            bw.writeInt(xf);
            bw.write("\"><v>");
            bw.writeInt(i);
            bw.write("</v></c>");
        } else {
            bw.write("\" t=\"inlineStr\" s=\"");
            bw.writeInt(xf);
            bw.write("\"><is><t>");
            bw.escapeWrite(s); // escape text
            bw.write("</t></is></c>");
        }

        // TODO optimize If auto-width
        if (hc.getAutoSize() == 1) {
            double ln;
            if (hc.o < (ln = stringWidth(s, xf))) hc.o = ln;
        }
    }

    /**
     * Write string value and cache the max string length
     *
     * @param s      the string value
     * @param row    the row index
     * @param column the column index
     * @param xf     the style index
     * @throws IOException if I/O error occur
     */
    @Deprecated
    protected void writeStringAutoSize(String s, int row, int column, int xf) throws IOException {
        writeString(s, row, column, xf);
//        Column hc = columns[column];
//        double ln;
//        if (hc.o < (ln = stringWidth(s, xf))) {
//            hc.o = ln;
//        }
    }

    /**
     * Write double value
     *
     * @param d      the double value
     * @param row    the row index
     * @param column the column index
     * @param xf     the style index
     * @throws IOException if I/O error occur
     */
    protected void writeDouble(double d, int row, int column, int xf) throws IOException {
        Column hc = columns[column];
        bw.write("<c r=\"");
        bw.write(int2Col(hc.getRealColIndex()));
        bw.writeInt(row);
        bw.write("\" s=\"");
        bw.writeInt(xf);
        bw.write("\"><v>");
        bw.write(d);
        bw.write("</v></c>");

        // TODO optimize If auto-width
        if (hc.getAutoSize() == 1) {
            int n;
            if (hc.o < (n = Double.toString(d).length())) hc.o = n;
        }
    }

    /**
     * Write double value and cache the max value
     *
     * @param d      the double value
     * @param row    the row index
     * @param column the column index
     * @param xf     the style index
     * @throws IOException if I/O error occur
     */
    @Deprecated
    protected void writeDoubleAutoSize(double d, int row, int column, int xf) throws IOException {
        writeDouble(d, row, column, xf);
//        Column hc = columns[column];
//        int n;
//        if (hc.width == 0 && hc.o < (n = Double.toString(d).length())) {
//            hc.o = n;
//        }
    }

    /**
     * Write decimal value
     *
     * @param bd     the decimal value
     * @param row    the row index
     * @param column the column index
     * @param xf     the style index
     * @throws IOException if I/O error occur
     */
    protected void writeDecimal(BigDecimal bd, int row, int column, int xf) throws IOException {
        Column hc = columns[column];
        bw.write("<c r=\"");
        bw.write(int2Col(hc.getRealColIndex()));
        bw.writeInt(row);
        bw.write("\" s=\"");
        bw.writeInt(xf);
        bw.write("\"><v>");
        bw.write(bd.toString());
        bw.write("</v></c>");
        // TODO optimize If auto-width
        if (hc.getAutoSize() == 1) {
            int l;
            if (hc.o < (l = bd.toString().length())) hc.o = l;
        }
    }

    /**
     * Write decimal value and cache the max value
     *
     * @param bd     the decimal value
     * @param row    the row index
     * @param column the column index
     * @param xf     the style index
     * @throws IOException if I/O error occur
     */
    @Deprecated
    protected void writeDecimalAutoSize(BigDecimal bd, int row, int column, int xf) throws IOException {
        writeDecimal(bd, row, column, xf);
//        Column hc = columns[column];
//        int l;
//        if (hc.width == 0 && hc.o < (l = bd.toString().length())) {
//            hc.o = l;
//        }
    }

    /**
     * Write char value
     *
     * @param c      the character value
     * @param row    the row index
     * @param column the column index
     * @param xf     the style index
     * @throws IOException if I/O error occur
     */
    protected void writeChar(char c, int row, int column, int xf) throws IOException {
        bw.write("<c r=\"");
        bw.write(int2Col(columns[column].getRealColIndex()));
        bw.writeInt(row);
        bw.write("\" t=\"s\" s=\"");
        bw.writeInt(xf);
        bw.write("\"><v>");
        bw.writeInt(sst.get(c));
        bw.write("</v></c>");
    }

    /**
     * Write numeric value
     *
     * @param l      the numeric value
     * @param row    the row index
     * @param column the column index
     * @param xf     the style index
     * @throws IOException if I/O error occur
     */
    protected void writeNumeric(long l, int row, int column, int xf) throws IOException {
        Column hc = columns[column];
        bw.write("<c r=\"");
        bw.write(int2Col(hc.getRealColIndex()));
        bw.writeInt(row);
        bw.write("\" s=\"");
        bw.writeInt(xf);
        bw.write("\"><v>");
        bw.write(l);
        bw.write("</v></c>");
        // TODO optimize If auto-width
        if (hc.getAutoSize() == 1) {
            int n;
            if (hc.o < (n = stringSize(l))) hc.o = n;
        }
    }

    /**
     * Write numeric value and cache the max value
     *
     * @param l      the numeric value
     * @param row    the row index
     * @param column the column index
     * @param xf     the style index
     * @throws IOException if I/O error occur
     */
    @Deprecated
    protected void writeNumericAutoSize(long l, int row, int column, int xf) throws IOException {
        writeNumeric(l, row, column, xf);
//        Column hc = columns[column];
//        int n;
//        if (hc.width == 0 && hc.o < (n = stringSize(l))) {
//            hc.o = n;
//        }
    }

    /**
     * Write boolean value
     *
     * @param bool   the boolean value
     * @param row    the row index
     * @param column the column index
     * @param xf     the style index
     * @throws IOException if I/O error occur
     */
    protected void writeBool(boolean bool, int row, int column, int xf) throws IOException {
        bw.write("<c r=\"");
        bw.write(int2Col(columns[column].getRealColIndex()));
        bw.writeInt(row);
        bw.write("\" t=\"b\" s=\"");
        bw.writeInt(xf);
        bw.write("\"><v>");
        bw.writeInt(bool ? 1 : 0);
        bw.write("</v></c>");
    }

    /**
     * Write blank value
     *
     * @param row    the row index
     * @param column the column index
     * @param xf     the style index
     * @throws IOException if I/O error occur
     */
    protected void writeNull(int row, int column, int xf) throws IOException {
        bw.write("<c r=\"");
        bw.write(int2Col(columns[column].getRealColIndex()));
        bw.writeInt(row);
        bw.write("\" s=\"");
        bw.writeInt(xf);
        bw.write("\"/>");
    }

    /**
     * Write binary file
     *
     * @param bytes  the binary data
     * @param row    the row index
     * @param column the column index
     * @param xf     the style index
     * @throws IOException if I/O error occur
     */
    protected void writeBinary(byte[] bytes, int row, int column, int xf) throws IOException {
        writeNull(row, column, xf);
        // Test file signatures
        FileSignatures.Signature signature = FileSignatures.test(ByteBuffer.wrap(bytes));
        if (signature == null) {
            LOGGER.warn("File types that are not allowed");
            return;
        }
        int id = sheet.getWorkbook().incrementMediaCounter();
        String name = "image" + id + "." + signature.extension;
        // Store in disk
        Files.write(mediaPath.resolve(name), bytes, StandardOpenOption.CREATE_NEW);

        // Write picture
        writePictureDirect(id, name, column, row, signature);
    }

    /**
     * Write file value
     *
     * @param path   the picture file
     * @param row    the row index
     * @param column the column index
     * @param xf     the style index
     * @throws IOException if I/O error occur
     */
    protected void writeFile(Path path, int row, int column, int xf) throws IOException {
        writeNull(row, column, xf);
        // Test file signatures
        FileSignatures.Signature signature = FileSignatures.test(path);
        if ("unknown".equals(signature.extension)) {
            LOGGER.warn("File types that are not allowed");
            return;
        }
        int id = sheet.getWorkbook().incrementMediaCounter();
        String name = "image" + id + "." + signature.extension;
        // Store
        Files.copy(path, mediaPath.resolve(name), StandardCopyOption.REPLACE_EXISTING);

        // Write picture
        writePictureDirect(id, name, column, row, signature);
    }

    /**
     * Write stream value
     *
     * @param stream  the picture input-stream
     * @param row    the row index
     * @param column the column index
     * @param xf     the style index
     * @throws IOException if I/O error occur
     */
    protected void writeStream(InputStream stream, int row, int column, int xf) throws IOException {
        writeNull(row, column, xf);

        byte[] bytes = new byte[1 << 13];
        int n;

        OutputStream os = null;
        try {
            n = stream.read(bytes);
            // Empty stream
            if (n <= 0) return;
            FileSignatures.Signature signature = FileSignatures.test(ByteBuffer.wrap(bytes, 0, n));
            if (signature == null) {
                LOGGER.warn("File types that are not allowed");
                return;
            }
            int id = sheet.getWorkbook().incrementMediaCounter();
            String name = "image" + id + "." + signature.extension;
            os = Files.newOutputStream(mediaPath.resolve(name));
            os.write(bytes, 0, n);

            if (n == bytes.length) {
                while ((n = stream.read(bytes)) > 0)
                    os.write(bytes, 0, n);
            }

            // Write picture
            writePictureDirect(id, name, column, row, signature);
        } catch (IOException ex) {
            LOGGER.warn("Copy stream error.", ex);
        } finally {
            try {
                stream.close();
            } catch (IOException e) { } // Ignore
            if (os != null) {
                try {
                    os.close();
                } catch (IOException e) { } // Ignore
            }
        }
    }

    /**
     * Write remote media value
     *
     * @param url  remote url
     * @param row    the row index
     * @param column the column index
     * @param xf     the style index
     * @throws IOException if I/O error occur
     */
    protected void writeRemoteMedia(String url, int row, int column, int xf) throws IOException {
        writeNull(row, column, xf);

        Picture picture = createPicture(column, row);
        picture.id = sheet.getWorkbook().incrementMediaCounter();

        // Async Drawing
        drawingsWriter.asyncDrawing(picture);

        // Supports asynchronous download
        downloadRemoteResource(picture, url);
    }

    /**
     * Resize column width
     *
     * @param path the sheet temp path
     * @param rows total of rows
     * @throws IOException if I/O error occur
     */
    protected void resizeColumnWidth(File path, int rows) throws IOException {
        // There has no column to reset width
        if (columns.length <= 0 || rows <= 0) return;
//        String[] widths = new String[columns.length];
        // Collect column width
        for (int i = 0; i < columns.length; i++) {
            Column hc = columns[i];
            int k = hc.getAutoSize();
            // If fixed width
            if (k == 2) {
                double width = hc.width >= 0.0D ? hc.width: sheet.getDefaultWidth();
//                widths[i] = BigDecimal.valueOf(Math.min(width + 0.65D, Const.Limit.COLUMN_WIDTH)).setScale(2, BigDecimal.ROUND_HALF_UP).toString();
                hc.width = BigDecimal.valueOf(Math.min(width + 0.65D, Const.Limit.COLUMN_WIDTH)).setScale(2, BigDecimal.ROUND_HALF_UP).doubleValue();
                continue;
            }
            double _l = stringWidth(hc.name, hc.getCellStyleIndex()), len;
            Class<?> clazz = hc.getClazz();
            if (isString(clazz)) {
                len = hc.o;
            }
            else if (isDateTime(clazz) || isLocalDateTime(clazz)) {
                if (hc.getNumFmt() != null) {
                    len = hc.getNumFmt().calcNumWidth(0);
                } else len = hc.o;
            }
            else if (isDate(clazz) || isLocalDate(clazz)) {
                if (hc.getNumFmt() != null) {
                    len = hc.getNumFmt().calcNumWidth(0);
                } else len = hc.o;
            }
            else if (isChar(clazz)) {
                len = 1;
            }
            else if (isInt(clazz) || isLong(clazz)) {
                // TODO Calculate character width based on numFmt
                if (hc.getNumFmt() != null) {
                    len = hc.getNumFmt().calcNumWidth(hc.o);
                } else len = hc.o;
            }
            else if (isFloat(clazz) || isDouble(clazz)) {
                // TODO Calculate character width based on numFmt
                if (hc.getNumFmt() != null) {
                    len = hc.getNumFmt().calcNumWidth(hc.o);
                } else len = hc.o;
            }
            else if (isBigDecimal(clazz)) {
                if (hc.getNumFmt() != null) {
                    len = hc.getNumFmt().calcNumWidth(hc.o);
                } else len = hc.o;
            }
            else if (isTime(clazz) || isLocalTime(clazz)) {
                if (hc.getNumFmt() != null) {
                    len = hc.getNumFmt().calcNumWidth(0);
                } else len = hc.o;
            }
            else if (isBool(clazz)) {
                len = 5.0D;
            }
            else if (hc.getNumFmt() != null) {
                len = hc.getNumFmt().calcNumWidth(0);
            }
            else {
                len = hc.o > 0 ? hc.o : 10.0D;
            }
            double width = Math.max(_l, len) + 1.86D;
            if (hc.width > 0.000001D) width = Math.min(width, hc.width + 0.65D);
            if (width > Const.Limit.COLUMN_WIDTH) {
                width = Const.Limit.COLUMN_WIDTH;
            }
//            widths[i] = BigDecimal.valueOf(width).setScale(2, BigDecimal.ROUND_HALF_UP).toString();
            hc.width = BigDecimal.valueOf(width).setScale(2, BigDecimal.ROUND_HALF_UP).doubleValue();
        }

        XMLWorksheetWriter _writer = new XMLWorksheetWriter(sheet);
        _writer.totalRows = totalRows;
        _writer.startRow = startRow;
        _writer.startHeaderRow = startHeaderRow;
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        _writer.bw = new ExtBufferedWriter(new OutputStreamWriter(baos, StandardCharsets.UTF_8));
        _writer.writeBefore();
        _writer.bw.close();

        // Resize each column width ...
        try (SeekableByteChannel channel = Files.newByteChannel(path.toPath(), StandardOpenOption.WRITE, StandardOpenOption.READ)) {
            ByteBuffer buffer = ByteBuffer.wrap(baos.toByteArray());
            buffer.order(ByteOrder.LITTLE_ENDIAN);
            channel.write(buffer);
        }
    }

    /**
     * Release resources
     */
    @Override
    public void close() {
        FileUtil.close(bw);
        // Close drawing writer
        FileUtil.close(drawingsWriter);
    }

    /**
     * Write the &lt;worksheet&gt; node
     *
     * @throws IOException if I/O error occur.
     */
    protected void writeRootNode() throws IOException {
        if (getClass().isAnnotationPresent(TopNS.class)) {
            TopNS topNS = getClass().getAnnotation(TopNS.class);
            bw.write('<');
            bw.write(topNS.value());
            String[] prefixs = topNS.prefix(), urls = topNS.uri();
            for (int i = 0, len = prefixs.length; i < len; ) {
                bw.write(" xmlns");
                if (prefixs[i] != null && !prefixs[i].isEmpty()) {
                    bw.write(':');
                    bw.write(prefixs[i]);
                }
                bw.write("=\"");
                bw.write(urls[i]);
                if (++i < len) {
                    bw.write('"');
                }
            }
        } else {
            bw.write("<worksheet xmlns=\"");
            bw.write(Const.SCHEMA_MAIN);
        }
        bw.write("\">");
    }

    /**
     * Write the dimension of sheet, default value is {@code A1}
     *
     * @throws IOException if I/O error occur.
     */
    protected void writeDimension() throws IOException {
        bw.append("<dimension ref=\"");
        if (columns.length > 0) bw.write(int2Col(columns[0].getRealColIndex()));
        else bw.write('A');
        bw.writeInt(startHeaderRow);
        int n = 11, size = sheet.size(); // fill 11 space
        if (size > 0 && columns.length > 0) {
            bw.write(':');
            n--;
            Column hc = columns[columns.length - 1];
            char[] col = int2Col(hc.getRealColIndex());
            bw.write(col);
            n -= col.length;
            bw.writeInt(size + startRow);
            n -= stringSize(size + startRow);
        }
        bw.write('"');
        for (; n-->0;) bw.write(32); // Fill space
        bw.write("/>");
    }

    /**
     * Write the sheet views such as FreezeEnum, Default selection cell.
     *
     * @throws IOException if I/O error occur.
     */
    protected void writeSheetViews() throws IOException {
        bw.write("<sheetViews>");

        bw.write("<sheetView workbookViewId=\"0\"");
        // Default show grid lines
        if (!sheet.isShowGridLines()) bw.write(" showGridLines=\"0\"");
        // Default select the first worksheet
        if (sheet.getId() == 1) bw.write(" tabSelected=\"1\"");

        // Freeze Panes
        Object o = sheet.getExtPropValue(Const.ExtendPropertyKey.FREEZE);
        if (o instanceof Panes) {
            Panes freezePanes = (Panes) o;

            // Validity check
            if (freezePanes.row < 0 || freezePanes.col < 0) {
                throw new IllegalArgumentException("Negative number occur in freeze panes settings.");
            }

            if ((freezePanes.col | freezePanes.row) == 0) {
                bw.write("/>"); // Empty tag
            } else {
                bw.write(">");

                Dimension dim = new Dimension(freezePanes.row + 1, (short) (freezePanes.col + 1));
                // Freeze top row
                if (freezePanes.col == 0) {
                    bw.write("<pane ySplit=\"" + freezePanes.row + "\" topLeftCell=\"" + dim + "\" activePane=\"bottomLeft\" state=\"frozen\"/>");
                    bw.write("<selection pane=\"bottomLeft\" activeCell=\"" + dim + "\" sqref=\"" + dim + "\"/>");
                }
                // Freeze first column
                else if (freezePanes.row == 0) {
                    bw.write("<pane xSplit=\"" + freezePanes.col + "\" topLeftCell=\"" + dim + "\" activePane=\"topRight\" state=\"frozen\"/>");
                    bw.write("<selection pane=\"topRight\" activeCell=\"" + dim + "\" sqref=\"" + dim + "\"/>");
                }
                // Freeze panes
                else {
                    bw.write("<pane xSplit=\"" + freezePanes.col + "\" ySplit=\"" + freezePanes.row + "\" topLeftCell=\"" + dim + "\" activePane=\"bottomRight\" state=\"frozen\"/>");
                    bw.write("<selection pane=\"topRight\" activeCell=\"" + new Dimension(1, dim.firstColumn) + "\" sqref=\"" + new Dimension(1, dim.firstColumn) + "\"/>");
                    bw.write("<selection pane=\"bottomLeft\" activeCell=\"" + new Dimension(dim.firstRow, (short) 1) + "\" sqref=\"" + new Dimension(dim.firstRow, (short) 1) + "\"/>");
                    bw.write("<selection pane=\"bottomRight\" activeCell=\"" + dim + "\" sqref=\"" + dim + "\"/>");
                }
                bw.write("</sheetView>");
            }
        }
        // Move the head row to the top
        else if (sheet.isScrollToVisibleArea() && startHeaderRow > 1) {
            bw.write(" topLeftCell=\"");
            char[] cols = int2Col(columns[0].realColIndex);
            bw.write(cols);
            bw.writeInt(startHeaderRow);
            bw.write("\">");
            bw.write("<selection activeCell=\"");
            bw.write(cols);
            bw.writeInt(startHeaderRow);
            bw.write("\" sqref=\"");
            bw.write(cols);
            bw.writeInt(startHeaderRow);
            bw.write("\"/></sheetView>");
        }
        else {
            bw.write("/>"); // Empty tag
        }
        bw.write("</sheetViews>");
    }

    /**
     * Write the sheet format
     *
     * @param fillSpace The number of characters to pad when recalculating the width.
     * @param defaultWidth The default cell width, {@code 8.38} will be use if it not be setting.
     * @throws IOException if I/O error occur.
     */
    protected void writeSheetFormat(int fillSpace, String defaultWidth) throws IOException {
        bw.write("<sheetFormatPr defaultRowHeight=\"18.5\" defaultColWidth=\"");
        bw.write(defaultWidth);
        bw.write('"');
        for (int i = fillSpace - defaultWidth.length(); i-->=0;) bw.write(32); // Fill space
        bw.write("/>");
    }

    /**
     * Write the default column info, The specified column width will be overwritten in these method.
     *
     * @param fillSpace The number of characters to pad when recalculating the width.
     * @param defaultWidth The default cell width, {@code 8.38} will be use if it not be setting.
     * @throws IOException if I/O error occur.
     */
    protected void writeCols(int fillSpace, String defaultWidth) throws IOException {
        if (columns.length > 0) {
            bw.write("<cols>");
            for (int i = 0; i < columns.length; i++) {
                Column col = columns[i];
                // Mark auto-width
                includeAutoWidth |= col.getAutoSize() == 1;
                String width = col.width >= 0.0000001D ? new BigDecimal(col.width).setScale(2, BigDecimal.ROUND_HALF_UP).toString() : defaultWidth;
                int w = width.length();
                bw.write("<col customWidth=\"1\" width=\"");
                bw.write(width);
                if (col.isHide()) {
                    bw.write("\" hidden=\"1");
                    w += 11;
                }
                bw.write('"');
                for (int j = fillSpace - w; j-- > 0; ) bw.write(32); // Fill space
                bw.write(" min=\"");
                bw.writeInt(col.getRealColIndex());
                bw.write("\" max=\"");
                bw.writeInt(col.getRealColIndex());
                bw.write("\" bestFit=\"1\"/>");
            }
            bw.write("</cols>");
        }
    }

    /**
     * Begin to write the Sheet data and the header row.
     *
     * @param nonHeader mark none header
     * @throws IOException if I/O error occur.
     */
    protected void beforeSheetData(boolean nonHeader) throws IOException {
        if (sheetDataReady > 0) return;
        // Start to write sheet data
        bw.write("<sheetData>");

        int headerRow = 0;
        // Write header rows
        if (!nonHeader && columns.length > 0) {
            headerRow = writeHeaderRow();
        }
        startRow = startHeaderRow + headerRow;
        sheetDataReady = 1;
    }

    /**
     * Append others customize info
     *
     * @throws IOException if I/O error occur.
     */
    protected void afterSheetData() throws IOException {
        // vmlDrawing
        Relationship r = sheet.findRel("vmlDrawing");
        if (r != null) {
            bw.write("<legacyDrawing r:id=\"");
            bw.write(r.getId());
            bw.write("\"/>");
        }

        // Compatible processing
        else if (comments != null) {
            sheet.addRel(r = new Relationship("../drawings/vmlDrawing" + sheet.getId() + Const.Suffix.VML, Const.Relationship.VMLDRAWING));
            sheet.addRel(new Relationship("../comments" + sheet.getId() + Const.Suffix.XML, Const.Relationship.COMMENTS));

            bw.write("<legacyDrawing r:id=\"");
            bw.write(r.getId());
            bw.write("\"/>");
        }

        // Background image
        if (sheet.getWaterMark() != null) {
            RelManager relManager = sheet.getRelManager();
            r = relManager.getByType(Const.Relationship.IMAGE);
            if (r != null) {
                bw.write("<picture r:id=\"");
                bw.write(r.getId());
                bw.write("\"/>");
            }
        }

        // Drawings
        if (drawingsWriter != null) {
            RelManager relManager = sheet.getRelManager();
            r = relManager.getByType(Const.Relationship.DRAWINGS);
            if (r != null) {
                bw.write("<drawing r:id=\"");
                bw.write(r.getId());
                bw.write("\"/>");
            }
        }
    }

    /**
     * Append merged cells if exists
     *
     * @throws IOException if I/O error occur.
     */
    protected void writeMergeCells() throws IOException {
        // Merge cells if exists
        @SuppressWarnings("unchecked")
        List<Dimension> mergeCells = (List<Dimension>) sheet.getExtPropValue(Const.ExtendPropertyKey.MERGE_CELLS);
        if (mergeCells != null && !mergeCells.isEmpty()) {
            bw.write("<mergeCells count=\"");
            bw.writeInt(mergeCells.size());
            bw.write("\">");
            for (Dimension dim : mergeCells) {
                bw.write("<mergeCell ref=\"");
                bw.write(dim.toString());
                bw.write("\"/>");
            }
            bw.write("</mergeCells>");
        }
    }

    /**
     * Char buffer
     */
    protected static char[] cacheChar = new char[1 << 8];

    /**
     * Calculate text width
     * FIXME reference {@link sun.swing.SwingUtilities2#stringWidth}
     *
     * @param s      the string value
     * @param xf     the style index
     * @return cell width
     */
    protected double stringWidth(String s, int xf) {
        if (StringUtil.isEmpty(s)) return 0.0D;
        int n = Math.min(s.length(), cacheChar.length);
        double w = 0.0D;
        s.getChars(0, n, cacheChar, 0);
        // TODO Calculate text width based on font-family and font-size
        for (int i = 0; i < n; w += cacheChar[i++] > 0x4E00 ? 1.86D : 1.0D);
        return w;
    }

    /**
     * Test whether the coordinate is the first cell of the merged cell,
     * and use {@link Grid#test} to determine yes before calling this method
     *
     * @param mergeCells all merged cells
     * @param row the cell row
     * @param col the cell column
     * @return true if the is first cell in merged
     */
    public static boolean isFirstMergedCell(List<Dimension> mergeCells, int row, int col) {
        for (Dimension dim : mergeCells) {
            if (dim.checkRange(row, col) && row == dim.firstRow && col == dim.firstColumn) return true;
        }
        return false;
    }

    /**
     * Returns the maximum cell height
     *
     * @param columnsArray the header column array
     * @param row actual rows in Excel
     * @return cell height or -1
     */
    public double getHeaderHeight(Column[][] columnsArray, int row) {
        double h = -1D;
        for (Column[] cols : columnsArray) h = Math.max(cols[row].headerHeight, h);
        return h;
    }

    /**
     * Create drawing writer and add relationship
     *
     * @return {@link IDrawingsWriter}
     */
    public IDrawingsWriter createDrawingsWriter() {
        int id = sheet.getWorkbook().incrementDrawingCounter();
        sheet.getWorkbook().addContentType(new ContentType.Override(Const.ContentType.DRAWINGS, "/xl/drawings/drawing" + id + ".xml"));
        sheet.addRel(new Relationship("../drawings/drawing" + id + ".xml", Const.Relationship.DRAWINGS));
        return new XMLDrawingsWriter(workSheetPath.getParent().resolve("drawings").resolve("drawing" + id + ".xml"));
    }

    /**
     * Download remote resources
     *
     * By default, only HTTP or HTTPS protocols are supported.
     * Use {@code HttpURLConnection} to synchronously download remote resources.
     * For more complex scenarios (asynchronous download, Connection pool, authentication, FTP, etc.), please override this method
     *
     * @param picture {@link Picture} info
     * @param url remote url
     */
    public void downloadRemoteResource(Picture picture, String url) throws IOException {
        try {
            // Support http or https
            if (url.charAt(0) == 'h') {
                HttpURLConnection con = (HttpURLConnection) new URL(url).openConnection();
                con.connect();
                InputStream is = con.getInputStream();
                if (is != null) {
                    ByteArrayOutputStream bos = new ByteArrayOutputStream(1 << 18);
                    int i;
                    byte[] bytes = new byte[1 << 18];
                    while ((i = is.read(bytes)) > 0) bos.write(bytes, 0, i);
                    downloadCompleted(picture, bos.toByteArray());
                    return;
                }
                con.disconnect();
            }
            // Ignore others
            downloadCompleted(picture, null);
        } catch (IOException e) {
            LOGGER.error("Download remote resource [{}] error", url, e);
            throw e;
        }
    }

    /**
     * Complete downloading, check the file signatures and set the subscript {@code Picture.idx} to idle
     *
     * @param picture {@link Picture} info
     * @param body thr resource data
     * @throws IOException if I/O error occur.
     */
    public void downloadCompleted(Picture picture, byte[] body) throws IOException {
        LOGGER.debug("completed Row: {} len: {}", picture.row, body != null ? body.length : 0);
        if (body == null || body.length == 0) {
            drawingsWriter.complete(picture);
            return;
        }
        // Test file signatures
        FileSignatures.Signature signature = FileSignatures.test(ByteBuffer.wrap(body));
        if (signature != null) {
            String name = "image" + picture.id + "." + signature.extension;
            // Store onto disk
            Files.write(mediaPath.resolve(name), body, StandardOpenOption.CREATE_NEW);
            picture.picName = name;
            picture.size = signature.width << 16 | signature.height;
            // Add global contentType
            sheet.getWorkbook().addContentType(new ContentType.Default(signature.contentType, signature.extension));
        } else LOGGER.warn("File types that are not allowed");

        // Setting idle
        drawingsWriter.complete(picture);
    }

    // Write picture
    protected void writePictureDirect(int id, String name, int column, int row, FileSignatures.Signature signature) throws IOException {
        Picture picture = createPicture(column, row);
        picture.id = id;
        picture.picName = name;
        picture.size = signature.width << 16 | signature.height;

        // Drawing
        drawingsWriter.drawing(picture);

        // Add global contentType
        sheet.getWorkbook().addContentType(new ContentType.Default(signature.contentType, signature.extension));
    }

    /**
     * Picture constructor
     * You can use this method to add general effects
     *
     * @param column cell column
     * @param row cell row
     * @return {@link Picture}
     */
    protected Picture createPicture(int column, int row) {
        Picture picture = new Picture();
        picture.col = column;
        picture.row = row;
        picture.padding = 1 << 24 | 1 << 16 | 1 << 8 | 1;

        return picture;
    }
}

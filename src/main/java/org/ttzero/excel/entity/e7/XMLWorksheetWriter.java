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

import org.ttzero.excel.annotation.TopNS;
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
import org.ttzero.excel.util.FileUtil;
import org.ttzero.excel.util.StringUtil;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.IOException;
import java.io.ObjectInputStream;
import java.io.ObjectOutputStream;
import java.lang.reflect.InvocationTargetException;
import java.math.BigDecimal;
import java.nio.ByteBuffer;
import java.nio.ByteOrder;
import java.nio.channels.SeekableByteChannel;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardOpenOption;
import java.util.List;
import java.util.function.Supplier;

import static org.ttzero.excel.entity.Sheet.int2Col;
import static org.ttzero.excel.reader.Cell.BLANK;
import static org.ttzero.excel.reader.Cell.BOOL;
import static org.ttzero.excel.reader.Cell.CHARACTER;
import static org.ttzero.excel.reader.Cell.DATE;
import static org.ttzero.excel.reader.Cell.DATETIME;
import static org.ttzero.excel.reader.Cell.DECIMAL;
import static org.ttzero.excel.reader.Cell.DOUBLE;
import static org.ttzero.excel.reader.Cell.INLINESTR;
import static org.ttzero.excel.reader.Cell.LONG;
import static org.ttzero.excel.reader.Cell.NUMERIC;
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
public class XMLWorksheetWriter implements IWorksheetWriter {

    // the storage path
    protected Path workSheetPath;
    protected ExtBufferedWriter bw;
    protected Sheet sheet;
    protected Column[] columns;
    protected final SharedStrings sst;
    protected Comments comments;
    protected int startRow;

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

        if (rowBlock != null && rowBlock.hasNext()) {
            if (sheet.isAutoSize()) {
                do {
                    // write row-block data auto size
                    writeAutoSizeRowBlock(rowBlock);
                    // end of row
                    if (rowBlock.isEOF()) break;
                } while ((rowBlock = supplier.get()) != null);
            } else {
                do {
                    // write row-block data
                    writeRowBlock(rowBlock);
                    // end of row
                    if (rowBlock.isEOF()) break;
                } while ((rowBlock = supplier.get()) != null);
            }
        }

        int total = rowBlock != null ? rowBlock.getTotal() : 0;

        // write end
        writeAfter(total);

        // Write some final info
        sheet.afterSheetAccess(workSheetPath);

        // resize
        if (sheet.isAutoSize()) {
            // close writer before resize
            close();
            resizeColumnWidth(sheetPath.toFile(), total);
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

        if (rowBlock.hasNext()) {
            if (sheet.isAutoSize()) {
                for (; ; ) {
                    // write row-block data
                    writeAutoSizeRowBlock(rowBlock);
                    // end of row
                    if (rowBlock.isEOF()) break;
                    // Get the next block
                    rowBlock = sheet.nextBlock();
                }
            } else {
                for (; ; ) {
                    // write row-block data
                    writeRowBlock(rowBlock);
                    // end of row
                    if (rowBlock.isEOF()) break;
                    // Get the next block
                    rowBlock = sheet.nextBlock();
                }
            }
        }

        // write end
        writeAfter(rowBlock.getTotal());

        // Write some final info
        sheet.afterSheetAccess(workSheetPath);

        // resize
        if (sheet.isAutoSize()) {
            // close writer before resize
            close();
            resizeColumnWidth(sheetPath.toFile(), rowBlock.getTotal());
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
        int fillSpace = 6;
        BigDecimal width = BigDecimal.valueOf(!nonHeader ? sheet.getDefaultWidth() : 8.38);
        String defaultWidth = width.setScale(2, BigDecimal.ROUND_HALF_UP).toString();
        writeSheetFormat(fillSpace, defaultWidth);

        // cols
        writeCols(fillSpace, defaultWidth);

        // Write body data
        beforeSheetDate(nonHeader);
    }

    /**
     * Write the header row
     *
     * @throws IOException if I/O error occur
     */
    protected int writeHeaderRow() throws IOException {
        // Write header
        int row = 0, subColumnSize = columns[0].subColumnSize(), defaultStyleIndex = sheet.defaultHeadStyleIndex();
        Column[][] columnsArray = new Column[columns.length][];
        for (int i = 0; i < columns.length; i++) {
            columnsArray[i] = columns[i].toArray();
        }
        // Merge cells if exists
        @SuppressWarnings("unchecked")
        List<Dimension> mergeCells = (List<Dimension>) sheet.getExtPropValue(Const.WorksheetExtendProperty.MERGE_CELLS);
        Grid mergedGrid = mergeCells != null && !mergeCells.isEmpty() ? GridFactory.create(mergeCells) : null;
        for (int i = subColumnSize - 1; i >= 0; i--) {

            row++;
            bw.write("<row r=\"");
            bw.writeInt(row);
            bw.write("\" customHeight=\"1\" ht=\"20.5\" spans=\"1:");
            bw.writeInt(columns[columns.length - 1].colIndex);
            bw.write("\">");

            String name;
            if (sheet.isAutoSize()) {
                for (int j = 0, c = 0; j < columns.length; j++) {
                    Column hc = columnsArray[j][i];
                    name = isNotEmpty(hc.getName()) ? hc.getName() : mergedGrid != null && mergedGrid.test(i + 1, j + 1) ? hc.name : hc.key;
                    writeStringAutoSize(name, row, c++, hc.getHeaderStyleIndex() == -1 ? defaultStyleIndex : hc.getHeaderStyleIndex());
                }
            } else {
                for (int j = 0, c = 0; j < columns.length; j++) {
                    Column hc = columnsArray[j][i];
                    name = isNotEmpty(hc.getName()) ? hc.getName() : mergedGrid != null && mergedGrid.test(i + 1, j + 1) ? hc.name : hc.key;
                    writeString(name, row, c++, hc.getHeaderStyleIndex() == -1 ? defaultStyleIndex : hc.getHeaderStyleIndex());
                }
            }

            // Write header comments
            for (int j = 0; j < columns.length; j++) {
                Column hc = columnsArray[j][i];
                if (hc.headerComment != null) {
                    if (comments == null) comments = sheet.createComments();
                    comments.addComment(new String(int2Col(hc.colIndex)) + row
                        , hc.headerComment.getTitle(), hc.headerComment.getValue());
                }
            }
            bw.write("</row>");
        }
        return row;
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

        // background image
        if (sheet.getWaterMark() != null) {
            // relationship
            Relationship r = sheet.findRel("media/image"); // only one background image
            if (r != null) {
                bw.write("<picture r:id=\"");
                bw.write(r.getId());
                bw.write("\"/>");
            }
        }
        // vmlDrawing
        Relationship r = sheet.findRel("vmlDrawing");
        if (r != null) {
            bw.write("<legacyDrawing r:id=\"");
            bw.write(r.getId());
            bw.write("\"/>");
        }

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
     */
    private void writeRowBlock(RowBlock rowBlock) throws IOException {
        for (; rowBlock.hasNext(); writeRow(rowBlock.next())) ;
    }

    /**
     * Write a row-block as auto size
     *
     * @param rowBlock the row-block
     */
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
     */
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
     * Write a row data
     *
     * @param row a row data
     * @throws IOException if I/O error occur
     */
    protected void writeRow(Row row) throws IOException {
        Cell[] cells = row.getCells();
        int len = cells.length, r = startRow(row.getIndex(), len);

        for (int i = 0; i < len; i++) {
            Cell cell = cells[i];
            int xf = cell.xf;
            switch (cell.t) {
                case INLINESTR:
                case SST:
                    writeString(cell.sv, r, i, xf);
                    break;
                case NUMERIC:
                    writeNumeric(cell.nv, r, i, xf);
                    break;
                case LONG:
                    writeNumeric(cell.lv, r, i, xf);
                    break;
                case DATE:
                case DATETIME:
                case DOUBLE:
                case TIME:
                    writeDouble(cell.dv, r, i, xf);
                    break;
                case BOOL:
                    writeBool(cell.bv, r, i, xf);
                    break;
                case DECIMAL:
                    writeDecimal(cell.mv, r, i, xf);
                    break;
                case CHARACTER:
                    writeChar(cell.cv, r, i, xf);
                    break;
                case BLANK:
                    writeNull(r, i, xf);
                    break;
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
    protected void writeRowAutoSize(Row row) throws IOException {
        Cell[] cells = row.getCells();
        int len = cells.length, r = startRow(row.getIndex(), len);

        for (int i = 0; i < len; i++) {
            Cell cell = cells[i];
            int xf = cell.xf;
            switch (cell.t) {
                case INLINESTR:
                case SST:
                    writeStringAutoSize(cell.sv, r, i, xf);
                    break;
                case NUMERIC:
                    writeNumericAutoSize(cell.nv, r, i, xf);
                    break;
                case LONG:
                    writeNumericAutoSize(cell.lv, r, i, xf);
                    break;
                case DATE:
                case DATETIME:
                case DOUBLE:
                case TIME:
                    writeDoubleAutoSize(cell.dv, r, i, xf);
                    break;
                case BOOL:
                    writeBool(cell.bv, r, i, xf);
                    break;
                case DECIMAL:
                    writeDecimalAutoSize(cell.mv, r, i, xf);
                    break;
                case CHARACTER:
                    writeChar(cell.cv, r, i, xf);
                    break;
                case BLANK:
                    writeNull(r, i, xf);
                    break;
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
        bw.write(int2Col(hc.colIndex));
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
    protected void writeStringAutoSize(String s, int row, int column, int xf) throws IOException {
        writeString(s, row, column, xf);
        Column hc = columns[column];
        int ln; // TODO get charset base on font style
        if (hc.width == 0 && hc.o < (ln = s.getBytes(StandardCharsets.UTF_8).length)) {
            hc.o = ln;
        }
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
        bw.write("<c r=\"");
        bw.write(int2Col(columns[column].colIndex));
        bw.writeInt(row);
        bw.write("\" s=\"");
        bw.writeInt(xf);
        bw.write("\"><v>");
        bw.write(d);
        bw.write("</v></c>");
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
    protected void writeDoubleAutoSize(double d, int row, int column, int xf) throws IOException {
        writeDouble(d, row, column, xf);
        Column hc = columns[column];
        int n;
        if (hc.width == 0 && hc.o < (n = Double.toString(d).length())) {
            hc.o = n;
        }
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
        bw.write("<c r=\"");
        bw.write(int2Col(columns[column].colIndex));
        bw.writeInt(row);
        bw.write("\" s=\"");
        bw.writeInt(xf);
        bw.write("\"><v>");
        bw.write(bd.toString());
        bw.write("</v></c>");
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
    protected void writeDecimalAutoSize(BigDecimal bd, int row, int column, int xf) throws IOException {
        writeDecimal(bd, row, column, xf);
        Column hc = columns[column];
        int l;
        if (hc.width == 0 && hc.o < (l = bd.toString().length())) {
            hc.o = l;
        }
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
        bw.write(int2Col(columns[column].colIndex));
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
        bw.write("<c r=\"");
        bw.write(int2Col(columns[column].colIndex));
        bw.writeInt(row);
        bw.write("\" s=\"");
        bw.writeInt(xf);
        bw.write("\"><v>");
        bw.write(l);
        bw.write("</v></c>");
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
    protected void writeNumericAutoSize(long l, int row, int column, int xf) throws IOException {
        writeNumeric(l, row, column, xf);
        Column hc = columns[column];
        int n;
        if (hc.width == 0 && hc.o < (n = (l < 0L ? stringSize(-l) + 1 : stringSize(l)))) {
            hc.o = n;
        }
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
        bw.write(int2Col(columns[column].colIndex));
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
        bw.write(int2Col(columns[column].colIndex));
        bw.writeInt(row);
        bw.write("\" s=\"");
        bw.writeInt(xf);
        bw.write("\"/>");
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
        if (columns.length <= 0 || rows <= 1) return;
        String[] widths = new String[columns.length];
        // Collect column width
        for (int i = 0; i < columns.length; i++) {
            Column hc = columns[i];
            double width = hc.width;
                // Fix width
            if (width < 0.0000001) {
                int _l = hc.name.getBytes(StandardCharsets.UTF_8).length, len;
                Class<?> clazz = hc.getClazz();
                // TODO Calculate text width based on font-family and font-size
                if (isString(clazz)) {
                    len = hc.o;
                }
                else if (isDate(clazz) || isLocalDate(clazz) || isDateTime(clazz) || isLocalDateTime(clazz)) {
//                    len = 10;
//                }
//                else if (isDateTime(clazz) || isLocalDateTime(clazz)) {
                    if (hc.getNumFmt() != null) {
                        len = hc.getNumFmt().calcNumWidth(0);
                    } else len = 20;
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
                    len = hc.o;
                }
                else if (isTime(clazz) || isLocalTime(clazz)) {
                    if (hc.getNumFmt() != null) {
                        len = hc.getNumFmt().calcNumWidth(0);
                    } else len = 8;
                }
                else if (isBool(clazz)) {
                    len = 5;
                }
                else {
                    len = 10;
                }
                width = _l > len ? _l + 3.38 : len + 3.38;
                if (width > Const.Limit.COLUMN_WIDTH) {
                    width = Const.Limit.COLUMN_WIDTH;
                }
            }
            widths[i] = BigDecimal.valueOf(width).setScale(2, BigDecimal.ROUND_HALF_UP).toString();
        }
        // resize each column width ...
        try (SeekableByteChannel channel = Files.newByteChannel(path.toPath(), StandardOpenOption.WRITE, StandardOpenOption.READ)) {
            ByteBuffer buffer = ByteBuffer.allocate(387 + columns.length * 73);
            buffer.order(ByteOrder.LITTLE_ENDIAN);

            int n = channel.read(buffer);
            if (n < 0) {
                throw new ExcelWriteException("Write worksheet [" + sheet.getName() + "] error.");
            }
            // Ready to read
            buffer.flip();

            // Rewrite dimension
            int position = findPosition(buffer, "<dimension ");
            // Get it
            if (position > 0) {
                buffer.put("ref=\"A1".getBytes(StandardCharsets.US_ASCII));
                int fill = 11; // fill 11 space
                buffer.put((byte) ':');
                fill--;
                char[] col = int2Col(columns[columns.length - 1].colIndex);
                buffer.put((new String(col) + (rows + 1)).getBytes(StandardCharsets.US_ASCII));
                fill -= col.length;
                fill -= stringSize(rows + 1);
                buffer.put((byte) '"');
                for (; fill-->0;) buffer.put((byte) 32); // Fill space
            }

            // Rewrite cols
            position = findPosition(buffer, "<cols>");
            if (position > 0) {
                for (String s : widths) {
                    position = findPosition(buffer, "width=\"");
                    if (position == -1) continue;
                    buffer.put(s.getBytes(StandardCharsets.US_ASCII));
                    buffer.put((byte) '"');
                    for (int j = 6 - s.length(); j-- > 0; ) buffer.put((byte) 32); // Fill space
                }
            }

            // Ready to write
            buffer.position(n);
            buffer.flip();
            // Move to header
            channel.position(0);
            channel.write(buffer);
        }
    }

    private int findPosition(ByteBuffer buffer, String key) {
        byte[] values = key.getBytes(StandardCharsets.UTF_8);
        for (; ; ) {
            for (; buffer.hasRemaining() && buffer.get() != values[0]; );
            if (!buffer.hasRemaining()) break;
            int j = 1;
            for (; j < values.length && buffer.hasRemaining() && buffer.get() == values[j++]; );
            if (j == values.length) {
                return 1;
            }
        }
        return -1;
    }

    /**
     * Release resources
     */
    @Override
    public void close() {
        FileUtil.close(bw);
    }

    /**
     * Write the &lt;worksheet&gt; node
     *
     * @throws IOException if I/O error occur.
     */
    protected void writeRootNode() throws IOException {
        if (sheet.getClass().isAnnotationPresent(TopNS.class)) {
            TopNS topNS = sheet.getClass().getAnnotation(TopNS.class);
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
        bw.append("<dimension ref=\"A1"); // FIXME Setting the column or row's start-index
        int n = 11, size = sheet.size(); // fill 11 space
        if (size > 0) {
            bw.write(':');
            n--;
            char[] col = int2Col(columns[columns.length - 1].colIndex);
            bw.write(col);
            n -= col.length;
            bw.writeInt(size + 1);
            n -= stringSize(size + 1);
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
        if (sheet.getId() == 1) { // Default select the first worksheet
            bw.write(" tabSelected=\"1\"");
        }

        Object o = sheet.getExtPropValue(Const.WorksheetExtendProperty.FREEZE);
        if (o instanceof Panes) {
            Panes freezePanes = (Panes) o;

            // Validity check
            if (freezePanes.row < 0 || freezePanes.col < 0) {
                throw new IllegalArgumentException("negative row or column number occur.");
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
        } else {
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
        bw.write("<sheetFormatPr defaultRowHeight=\"15.5\" defaultColWidth=\"");
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
//            if (sheet.isAutoSize()) {
            for (int i = 0; i < columns.length; i++) {
                Column col = columns[i];
                bw.write("<col customWidth=\"1\" width=\"");
                bw.write(col.width > 0.0000001 ? new BigDecimal(col.width).setScale(2, BigDecimal.ROUND_HALF_UP).toString() : defaultWidth);
                bw.write('"');
                for (int j = fillSpace - defaultWidth.length(); j-- > 0; ) bw.write(32); // Fill space
                bw.write(" min=\"");
                bw.writeInt(col.colIndex);
                bw.write("\" max=\"");
                bw.writeInt(col.colIndex + 1);
                bw.write("\" bestFit=\"1\"/>");
            }
//            } else {
//                bw.write("<col customWidth=\"1\" width=\"");
//                bw.write(columns[i].width >= 1.0 ? new BigDecimal(columns[i].width).setScale(2, BigDecimal.ROUND_HALF_UP).toString() : stringWidth);
//                bw.write('"');
//                for (int j = n; j-- > 0; ) bw.write(32); // Fill space
//                bw.write(" min=\"");
//                bw.writeInt(columns[0].colIndex);
//                bw.write("\" max=\"");
//                bw.writeInt(columns[columns.length - 1].colIndex + 1);
//                bw.write("\" bestFit=\"1\"/>");
//            }
            bw.write("</cols>");
        }
    }

    /**
     * Begin to write the Sheet data and the header row.
     *
     * @param nonHeader mark none header
     * @throws IOException if I/O error occur.
     */
    protected void beforeSheetDate(boolean nonHeader) throws IOException {
        bw.write("<sheetData>");

        int headerRow = 1;
        // The first header row
        if (!nonHeader) {
            headerRow = writeHeaderRow();
        }
        startRow = headerRow + (sheet.getNonHeader() ^ 1);
    }

    /**
     * Append others customize info
     *
     * @throws IOException if I/O error occur.
     */
    protected void afterSheetData() throws IOException {
        // Merge cells if exists
        @SuppressWarnings("unchecked")
        List<Dimension> mergeCells = (List<Dimension>) sheet.getExtPropValue(Const.WorksheetExtendProperty.MERGE_CELLS);
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
}

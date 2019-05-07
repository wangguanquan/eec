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

package cn.ttzero.excel.entity.e7;

import cn.ttzero.excel.annotation.TopNS;
import cn.ttzero.excel.entity.*;
import cn.ttzero.excel.manager.Const;
import cn.ttzero.excel.reader.Cell;
import cn.ttzero.excel.util.ExtBufferedWriter;
import cn.ttzero.excel.util.FileUtil;
import cn.ttzero.excel.util.StringUtil;

import java.io.*;
import java.lang.reflect.InvocationTargetException;
import java.math.BigDecimal;
import java.nio.ByteBuffer;
import java.nio.channels.FileChannel;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.util.function.Supplier;

import static cn.ttzero.excel.reader.Cell.*;
import static cn.ttzero.excel.entity.IWorksheetWriter.*;
import static cn.ttzero.excel.util.ExtBufferedWriter.stringSize;

/**
 * Create by guanquan.wang at 2019-04-22 16:31
 */
public class XMLWorksheetWriter implements IWorksheetWriter {

    private int headInfoLen, baseInfoLen;
    // the storage path
    private Path workSheetPath;
    private ExtBufferedWriter bw;
    private Sheet sheet;
    private Sheet.Column[] columns;
    private SharedStrings sst;

    public XMLWorksheetWriter(Sheet sheet) {
        this.sheet = sheet;
        this.sst = sheet.getSst();
    }

    /**
     * Write a row block
     *
     * @param supplier a row-block supplier
     * @throws IOException if io error occur
     */
    @Override
    public void write(Path path, Supplier<RowBlock> supplier) throws IOException {
        this.workSheetPath = path.resolve("worksheets");
        if (!Files.exists(this.workSheetPath)) {
            FileUtil.mkdir(workSheetPath);
        }
        sheet.what("0010", sheet.getName());

        Path sheetPath = workSheetPath.resolve(sheet.getFileName());

        this.bw = new ExtBufferedWriter(Files.newBufferedWriter(
            sheetPath, StandardCharsets.UTF_8));

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
                    if (rowBlock.isEof()) break;
                } while ((rowBlock = supplier.get()) != null);
            } else {
                do {
                    // write row-block data
                    writeRowBlock(rowBlock);
                    // end of row
                    if (rowBlock.isEof()) break;
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
     * @throws IOException if io error occur
     */
    @Override
    public void write(Path path) throws IOException {
        this.workSheetPath = path.resolve("worksheets");
        if (!Files.exists(this.workSheetPath)) {
            FileUtil.mkdir(workSheetPath);
        }
        sheet.what("0010", sheet.getName());

        Path sheetPath = workSheetPath.resolve(sheet.getFileName());

        this.bw = new ExtBufferedWriter(Files.newBufferedWriter(
            sheetPath, StandardCharsets.UTF_8));

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
                    if (rowBlock.isEof()) break;
                    // Get the next block
                    rowBlock = sheet.nextBlock();
                }
            } else {
                for (; ; ) {
                    // write row-block data
                    writeRowBlock(rowBlock);
                    // end of row
                    if (rowBlock.isEof()) break;
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
        IWorksheetWriter copy = null;
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
                    e2.printStackTrace();
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
     */
    protected void writeBefore() throws IOException {
        // The header columns
        columns = sheet.getHeaderColumns();
        boolean noneHeader = columns == null || columns.length == 0;

        StringBuilder buf = new StringBuilder(Const.EXCEL_XML_DECLARATION);
        // Declaration
        buf.append(Const.lineSeparator); // new line
        // Root node
        if (sheet.getClass().isAnnotationPresent(TopNS.class)) {
            TopNS topNS = sheet.getClass().getAnnotation(TopNS.class);
            buf.append('<').append(topNS.value());
            String[] prefixs = topNS.prefix(), urls = topNS.uri();
            for (int i = 0, len = prefixs.length; i < len; ) {
                buf.append(" xmlns");
                if (prefixs[i] != null && !prefixs[i].isEmpty()) {
                    buf.append(':').append(prefixs[i]);
                }
                buf.append("=\"").append(urls[i]);
                if (++i < len) {
                    buf.append('"');
                }
            }
        } else {
            buf.append("<worksheet xmlns=\"").append(Const.SCHEMA_MAIN);
        }
        buf.append("\">");

        // Dimension
        buf.append("<dimension ref=\"A1");
//        if (sheet.size() > 0) {
//            buf.append(":").append(sheet.int2Col(columns.length)).append(sheet.size() + 1);
//        }
        buf.append("\"/>");
        headInfoLen = buf.length() - 3;

        // SheetViews default value
        buf.append("<sheetViews><sheetView workbookViewId=\"0\"");
        if (sheet.getId() == 1) { // Default select the first worksheet
            buf.append(" tabSelected=\"1\"");
        }
        buf.append("/></sheetViews>");

        // Default format
        // throw unknown tag if not auto-size
        if (!noneHeader) {
            buf.append("<sheetFormatPr defaultRowHeight=\"16.5\" defaultColWidth=\"");
            buf.append(sheet.getDefaultWidth());
            buf.append("\"/>");
        } else {
            buf.append("<sheetFormatPr defaultRowHeight=\"13.5\" defaultColWidth=\"8.38\" />");
        }

        baseInfoLen = buf.length() - headInfoLen;
        // Write base info
        bw.write(buf.toString());


//        // cols
//        bw.write("<cols>");
//        for (int i = 0; i < columns.length; i++) {
//            bw.write("<col customWidth=\"1\" width=\"");
//            bw.write(sheet.getDefaultWidth());
//            bw.write("\" max=\"");
//            bw.writeInt(i + 1);
//            bw.write("\" min=\"");
//            bw.writeInt(i + 1);
//            bw.write("\" bestFit=\"1\"/>");
//        }
//        bw.write("</cols>");

        // Write body data
        bw.write("<sheetData>");

        if (!noneHeader) {
            writeHeaderRow();
        }
    }

    /**
     * Write the header row
     *
     * @throws IOException if io error occur
     */
    protected void writeHeaderRow() throws IOException {
        // Write header
        int row = 1;
        bw.write("<row r=\"");
        bw.writeInt(row);
        bw.write("\" customHeight=\"1\" ht=\"18.6\" spans=\"1:");
        bw.writeInt(columns.length);
        bw.write("\">");

        int c = 1, defaultStyle = sheet.defaultHeadStyle();
        for (Sheet.Column hc : columns) {
            bw.write("<c r=\"");
            bw.write(sheet.int2Col(c++));
            bw.writeInt(row);
            bw.write("\" t=\"inlineStr\" s=\"");
            bw.writeInt(defaultStyle);
            bw.write("\"><is><t>");
            bw.write(hc.getName());
            bw.write("</t></is></c>");
        }
        bw.write("</row>");
    }

    /**
     * 写尾部
     */
    protected void writeAfter(int total) throws IOException {
        // End target --sheetData
        bw.write("</sheetData>");

        // background image
        if (sheet.getWaterMark() != null) {
            // relationship
            Relationship r = sheet.find("media/image"); // only one background image
            if (r != null) {
                bw.write("<picture r:id=\"");
                bw.write(r.getId());
                bw.write("\"/>");
            }
        }
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
     * @throws IOException
     */
    protected int startRow(int rows, int columns) throws IOException {
        // Row number
        int r = rows + 2;
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
                case DATE:
                case NUMERIC:
                    writeNumeric(cell.nv, r, i, xf);
                    break;
                case LONG:
                    writeNumeric(cell.lv, r, i, xf);
                    break;
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
                case DATE:
                case NUMERIC:
                    writeNumericAutoSize(cell.nv, r, i, xf);
                    break;
                case LONG:
                    writeNumericAutoSize(cell.lv, r, i, xf);
                    break;
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
     * @throws IOException if io error occur
     */
    protected void writeString(String s, int row, int column, int xf) throws IOException {
        Sheet.Column hc = columns[column];
        bw.write("<c r=\"");
        bw.write(sheet.int2Col(column + 1));
        bw.writeInt(row);
        int i;
        if (StringUtil.isEmpty(s)) {
            bw.write("\" s=\"");
            bw.writeInt(xf);
            bw.write("\"/>");
        }
        // FIXME default string value is shared
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
     * @throws IOException if io error occur
     */
    protected void writeStringAutoSize(String s, int row, int column, int xf) throws IOException {
        writeString(s, row, column, xf);
        Sheet.Column hc = columns[column];
        int ln; // TODO get charset base on font style
        if (hc.width == 0 && hc.o < (ln = s.getBytes("GBK").length)) {
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
     * @throws IOException if io error occur
     */
    protected void writeDouble(double d, int row, int column, int xf) throws IOException {
        bw.write("<c r=\"");
        bw.write(sheet.int2Col(column + 1));
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
     * @throws IOException if io error occur
     */
    protected void writeDoubleAutoSize(double d, int row, int column, int xf) throws IOException {
        writeDouble(d, row, column, xf);
        Sheet.Column hc = columns[column];
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
     * @throws IOException if io error occur
     */
    protected void writeDecimal(BigDecimal bd, int row, int column, int xf) throws IOException {
        bw.write("<c r=\"");
        bw.write(sheet.int2Col(column + 1));
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
     * @throws IOException if io error occur
     */
    protected void writeDecimalAutoSize(BigDecimal bd, int row, int column, int xf) throws IOException {
        writeDecimal(bd, row, column, xf);
        Sheet.Column hc = columns[column];
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
     * @throws IOException if io error occur
     */
    protected void writeChar(char c, int row, int column, int xf) throws IOException {
        bw.write("<c r=\"");
        bw.write(sheet.int2Col(column + 1));
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
     * @throws IOException if io error occur
     */
    protected void writeNumeric(long l, int row, int column, int xf) throws IOException {
        bw.write("<c r=\"");
        bw.write(sheet.int2Col(column + 1));
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
     * @throws IOException if io error occur
     */
    protected void writeNumericAutoSize(long l, int row, int column, int xf) throws IOException {
        writeNumeric(l, row, column, xf);
        Sheet.Column hc = columns[column];
        int n;
        if (hc.width == 0 && hc.o < (n = stringSize(l))) {
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
     * @throws IOException if io error occur
     */
    protected void writeBool(boolean bool, int row, int column, int xf) throws IOException {
        bw.write("<c r=\"");
        bw.write(sheet.int2Col(column + 1));
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
     * @throws IOException if io error occur
     */
    protected void writeNull(int row, int column, int xf) throws IOException {
        bw.write("<c r=\"");
        bw.write(sheet.int2Col(column + 1));
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
     * @throws IOException if io error occur
     */
    protected void resizeColumnWidth(File path, int rows) throws IOException {
        // There has no column to reset width
        if (columns.length <= 0 || rows <= 1) return;
        // resize each column width ...
        File temp = new File(path.getParent(), sheet.getName() + ".temp");
        if (!path.renameTo(temp)) {
            Files.move(path.toPath(), temp.toPath(), StandardCopyOption.REPLACE_EXISTING);
        }

        FileChannel inChannel = null;
        FileChannel outChannel = null;
        try (FileInputStream fis = new FileInputStream(temp);
             FileOutputStream fos = new FileOutputStream(path)) {
            inChannel = fis.getChannel();
            outChannel = fos.getChannel();

            inChannel.transferTo(0, headInfoLen, outChannel);
            ByteBuffer buffer = ByteBuffer.allocate(baseInfoLen);
            inChannel.read(buffer, headInfoLen);
            buffer.compact();
            byte b;
            if ((b = buffer.get()) == '"') {
                char[] chars = sheet.int2Col(columns.length > 0 ? columns.length : 1);
                String s = ':' + new String(chars) + (rows + 1);
                outChannel.write(ByteBuffer.wrap(s.getBytes(StandardCharsets.UTF_8)));
            }
            buffer.flip();
            buffer.put(b);
            buffer.compact();
            outChannel.write(buffer);

            StringBuilder buf = new StringBuilder();
            buf.append("<cols>");
            int i = 0;
            for (Sheet.Column hc : columns) {
                i++;
                buf.append("<col customWidth=\"1\" width=\"");
                // Fix width
                if (hc.width > 0.0000001) {
                    buf.append(hc.width);
                    buf.append("\" max=\"");
                    buf.append(i);
                    buf.append("\" min=\"");
                    buf.append(i);
                    buf.append("\"/>");
                } else {
                    int _l = hc.name.getBytes("GB2312").length, len;
                    Class<?> clazz = hc.getClazz();
                    // TODO Calculate text width based on font-family and font-size
                    if (isString(clazz)) {
                        len = hc.o;
                    }
                    else if (isDate(clazz) || isLocalDate(clazz)) {
                        len = 10;
                    }
                    else if (isDateTime(clazz) || isLocalDateTime(clazz)) {
                        len = 20;
                    }
                    else if (isChar(clazz)) {
                        len = 1;
                    }
                    else if (isInt(clazz) || isLong(clazz)) {
                        // TODO Calculate character width based on numFmt
                        len = hc.o;
                    }
                    else if (isFloat(clazz) || isDouble(clazz)) {
                        // TODO Calculate character width based on numFmt
                        len = hc.o;
                    }
                    else if (isBigDecimal(clazz)) {
                        len = hc.o;
                    }
                    else if (isTime(clazz) || isLocalTime(clazz)) {
                        len = 8;
                    }
                    else if (isBool(clazz)) {
                        len = 5;
                    }
                    else {
                        len = 10;
                    }
                    buf.append(_l > len ? _l + 3.38 : len + 3.38);
                    buf.append("\" max=\"");
                    buf.append(i);
                    buf.append("\" min=\"");
                    buf.append(i);
                    buf.append("\" bestFit=\"1\"/>");
                }
            }
            buf.append("</cols>");

            outChannel.write(ByteBuffer.wrap(buf.toString().getBytes(StandardCharsets.UTF_8)));
            int start = headInfoLen + baseInfoLen;
            inChannel.transferTo(start, inChannel.size() - start, outChannel);

        } finally {
            boolean delete = temp.delete();
            if (!delete) {
                sheet.what("9005", temp.getAbsolutePath());
            }
            if (inChannel != null) {
                inChannel.close();
            }
            if (outChannel != null) {
                outChannel.close();
            }
        }
    }

    /**
     * Release resources
     */
    @Override
    public void close() {
        FileUtil.close(bw);
    }
}

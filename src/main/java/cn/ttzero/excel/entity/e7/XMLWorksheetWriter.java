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
import cn.ttzero.excel.entity.style.Styles;
import cn.ttzero.excel.manager.Const;
import cn.ttzero.excel.reader.Cell;
import cn.ttzero.excel.util.ExtBufferedWriter;
import cn.ttzero.excel.util.FileUtil;
import cn.ttzero.excel.util.StringUtil;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.ref.WeakReference;
import java.math.BigDecimal;
import java.nio.ByteBuffer;
import java.nio.channels.FileChannel;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Time;
import java.sql.Timestamp;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.util.Date;
import java.util.function.Supplier;

import static cn.ttzero.excel.reader.Cell.*;
import static cn.ttzero.excel.util.DateUtil.toDateTimeValue;
import static cn.ttzero.excel.util.DateUtil.toDateValue;
import static cn.ttzero.excel.util.DateUtil.toTimeValue;

/**
 * Create by guanquan.wang at 2019-04-22 16:31
 */
public class XMLWorksheetWriter implements IWorksheetWriter {

    private int headInfoLen, baseInfoLen;
    // the storage path
    private Path workSheetPath;
//    private Sheet sheet;
    private Workbook workbook;
    private ExtBufferedWriter bw;
    private Sheet sheet;
    private Sheet.Column[] columns;
    private SharedStrings sst;

    public XMLWorksheetWriter(Workbook workbook, Sheet sheet) {
//        this.sheet = sheet;
        this.workbook = workbook;
        this.sst = workbook.getSst();
        this.sheet = sheet;
    }

    /**
     * Write a row block
     * @param supplier a row-block supplier
     * @throws IOException if io error occur
     */
    @Override
    public void write(Path path, Supplier<RowBlock> supplier) throws IOException {
        this.workSheetPath = path.resolve("worksheets");
        if (!Files.exists(this.workSheetPath)) {
            FileUtil.mkdir(workSheetPath);
        }
        workbook.what("0010", sheet.getName());

        this.bw = new ExtBufferedWriter(Files.newBufferedWriter(
                workSheetPath.resolve(sheet.getFileName()), StandardCharsets.UTF_8));

        // write before
        writeBefore();

        RowBlock rowBlock;
        while ((rowBlock = supplier.get()) != null) {
            // write row-block data
            for (; rowBlock.hasNext(); writeRow(rowBlock.next()));
            // end of row
            if (rowBlock.isEof()) break;
        }

        // write end
        writeAfter(rowBlock != null ? rowBlock.getTotal() : 0);

        // TODO resize

        // Write some final info
        sheet.afterSheetAccess(workSheetPath);
    }

    /**
     * Write a row block
     * @param path the storage path
     * @throws IOException if io error occur
     */
    @Override
    public void write(Path path) throws IOException {
        this.workSheetPath = path.resolve("worksheets");
        if (!Files.exists(this.workSheetPath)) {
            FileUtil.mkdir(workSheetPath);
        }
        workbook.what("0010", sheet.getName());

        this.bw = new ExtBufferedWriter(Files.newBufferedWriter(
            workSheetPath.resolve(sheet.getFileName()), StandardCharsets.UTF_8));

        // write before
        writeBefore();

        RowBlock rowBlock;
        for ( ; ; ) {
            rowBlock = sheet.nextBlock();
            // write row-block data
            for (; rowBlock.hasNext(); writeRow(rowBlock.next()));
            // end of row
            if (rowBlock.isEof()) break;
        }

        // write end
        writeAfter(rowBlock.getTotal());

        // TODO resize

        // Write some final info
        sheet.afterSheetAccess(workSheetPath);
    }

    @Override
    public IWorksheetWriter copy(Sheet sheet) {
        return new XMLWorksheetWriter(workbook, sheet);
    }

    /**
     * The Worksheet row limit
     * @return the limit
     */
    @Override
    public int getRowLimit() {
        return Const.Limit.MAX_ROWS_ON_SHEET;
    }

    /**
     * The Worksheet column limit
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
        if (sheet.size() > 0) {
            buf.append(":").append(sheet.int2Col(columns.length)).append(sheet.size() + 1);
        }
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
        buf.append("<sheetFormatPr defaultRowHeight=\"16.5\" defaultColWidth=\"");
        buf.append(sheet.getDefaultWidth());
        buf.append("\"/>");

        baseInfoLen = buf.length() - headInfoLen;
        // Write base info
        bw.write(buf.toString());


        // cols
        bw.write("<cols>");
        for (int i = 0; i < columns.length; i++) {
            bw.write("<col customWidth=\"1\" width=\"");
            bw.write(sheet.getDefaultWidth());
            bw.write("\" max=\"");
            bw.writeInt(i + 1);
            bw.write("\" min=\"");
            bw.writeInt(i + 1);
            bw.write("\" bestFit=\"1\"/>");
        }
        bw.write("</cols>");

        // Write body data
        bw.write("<sheetData>");

        // Write header
        int r = 1;
        bw.write("<row r=\"");
        bw.writeInt(r);
        bw.write("\" customHeight=\"1\" ht=\"18.6\" spans=\"1:"); // spans 指定row开始和结束行
        bw.writeInt(columns.length);
        bw.write("\">");

        int c = 1, defaultStyle = sheet.defaultHeadStyle();
        for (Sheet.Column hc : columns) {
            bw.write("<c r=\"");
            bw.write(sheet.int2Col(c++));
            bw.writeInt(r);
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
        if (workbook.getWaterMark() != null) {
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
        workbook.what("0009", sheet.getName(), String.valueOf(total));
    }

    /**
     * Write a row-block
     * @param rowBlock the row-block
     */
    void writeRowBlock(RowBlock rowBlock) throws IOException {
        for (; rowBlock.hasNext(); writeRow(rowBlock.next()));
    }

    /**
     * Write a row data
     * @param row a row data
     */
    private void writeRow(Row row) throws IOException {
        // Row number
        int r = row.getIndex() + 2;
        // logging
        if (r % 1_0000 == 0) {
            workbook.what("0014", String.valueOf(r));
        }
        Cell[] cells = row.getCells();
        int len = cells.length;
        bw.write("<row r=\"");
        bw.writeInt(r);
        // default data row height 16.5
        bw.write("\" spans=\"1:");
        bw.writeInt(len);
        bw.write("\">");

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

//    /**
//     * 写行数据
//     * @param rs ResultSet
//     * @param bw bufferedWriter
//     */
//    protected void writeRow(ResultSet rs, ExtBufferedWriter bw) throws IOException, SQLException {
//        // Row number
//        int r = ++rows;
//        // logging
//        if (r % 1_0000 == 0) {
//            workbook.what("0014", String.valueOf(r));
//        }
//        final int len = columns.length;
//        bw.write("<row r=\"");
//        bw.writeInt(r);
//        // default data row height 16.5
//        bw.write("\" spans=\"1:");
//        bw.writeInt(len);
//        bw.write("\">");
//
//        for (int i = 0; i < len; i++) {
//            Column hc = columns[i];
//
//            // t n=numeric (default), s=string, b=boolean, str=function string
//            // TODO function <f ca="1" or t="shared" ref="O10:O15" si="0" ... si="10"></f>
//            if (isString(hc.clazz)) {
//                String s = rs.getString(i + 1);
//                writeString(bw, s, i);
//            }
//            else if (isDate(hc.clazz)) {
//                java.sql.Date date = rs.getDate(i + 1);
//                writeDate(bw, date, i);
//            }
//            else if (isDateTime(hc.clazz)) {
//                Timestamp ts = rs.getTimestamp(i + 1);
//                writeTimestamp(bw, ts, i);
//            }
////            else if (isChar(hc.clazz)) {
////                char c = (char) rs.getInt(i + 1);
////                writeChar(bw, c, i);
////            }
//            else if (isInt(hc.clazz)) {
//                int n = rs.getInt(i + 1);
//                writeInt(bw, n, i);
//            }
//            else if (isLong(hc.clazz)) {
//                long l = rs.getLong(i + 1);
//                writeLong(bw, l, i);
//            }
//            else if (isFloat(hc.clazz)) {
//                double d = rs.getDouble(i + 1);
//                writeDouble(bw, d, i);
//            } else if (isBool(hc.clazz)) {
//                boolean bool = rs.getBoolean(i + 1);
//                writeBoolean(bw, bool, i);
//            } else if (isBigDecimal(hc.clazz)) {
//                writeBigDecimal(bw, rs.getBigDecimal(i + 1), i);
//            } else if (isTime(hc.clazz)) {
//                writeTime(bw, rs.getTime(i + 1), i);
//            } else {
//                Object o = rs.getObject(i + 1);
//                if (o != null) {
//                    writeString(bw, o.toString(), i);
//                } else {
//                    writeNull(bw, i);
//                }
//            }
//        }
//        bw.write("</row>");
//    }
//
//    /**
//     * 写行数据
//     * @param rs ResultSet
//     * @param bw
//     */
//    protected void writeRowAutoSize(ResultSet rs, ExtBufferedWriter bw) throws IOException, SQLException {
//        int r = ++rows;
//        // logging
//        if (r % 1_0000 == 0) {
//            workbook.what("0014", String.valueOf(r));
//        }
//        final int len = columns.length;
//        bw.write("<row r=\"");
//        bw.writeInt(r);
//        bw.write("\" spans=\"1:");
//        bw.writeInt(len);
//        bw.write("\">");
//
//        for (int i = 0; i < len; i++) {
//            Column hc = columns[i];
//            // t n=numeric (default), s=string, b=boolean, str=function string
//            // TODO function <f ca="1" or t="shared" ref="O10:O15" si="0" ... si="10"></f>
//            if (isString(hc.clazz)) {
//                String s = rs.getString(i + 1);
//                writeStringAutoSize(bw, s, i);
//            }
//            else if (isDate(hc.clazz)) {
//                java.sql.Date date = rs.getDate(i + 1);
//                writeDate(bw, date, i);
//            }
//            else if (isDateTime(hc.clazz)) {
//                Timestamp ts = rs.getTimestamp(i + 1);
//                writeTimestamp(bw, ts, i);
//            }
//            else if (isInt(hc.clazz)) {
//                int n = rs.getInt(i + 1);
//                writeIntAutoSize(bw, n, i);
//            }
//            else if (isLong(hc.clazz)) {
//                long l = rs.getLong(i + 1);
//                writeLong(bw, l, i);
//            }
//            else if (isFloat(hc.clazz)) {
//                double d = rs.getDouble(i + 1);
//                writeDouble(bw, d, i);
//            } else if (isBool(hc.clazz)) {
//                boolean bool = rs.getBoolean(i + 1);
//                writeBoolean(bw, bool, i);
//            } else if (isBigDecimal(hc.clazz)) {
//                writeBigDecimal(bw, rs.getBigDecimal(i + 1), i);
//            } else if (isTime(hc.clazz)) {
//                writeTime(bw, rs.getTime(i + 1), i);
//            } else {
//                Object o = rs.getObject(i + 1);
//                if (o != null) {
//                    writeStringAutoSize(bw, o.toString(), i);
//                } else {
//                    writeNull(bw, i);
//                }
//            }
//        }
//        bw.write("</row>");
//    }
//




//    /**
//     * 写字符串
//     * @throws IOException
//     */
//    protected void writeString(String s, int row, int column) throws IOException {
//        writeString(s, row, column, s);
//    }

    private void writeString(String s, int row, int column, int xf) throws IOException {
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
        }
        else {
            bw.write("\" t=\"inlineStr\" s=\"");
            bw.writeInt(xf);
            bw.write("\"><is><t>");
            bw.escapeWrite(s); // escape text
            bw.write("</t></is></c>");
        }
    }

//    protected void writeStringAutoSize(ExtBufferedWriter bw, String s, int column) throws IOException {
//        writeStringAutoSize(bw, s, column, s);
//    }
//
//    protected void writeStringAutoSize(ExtBufferedWriter bw, String s, int column, Object o) throws IOException {
//        Column hc = columns[column];
//        int styleIndex = getStyleIndex(hc, o);
//        bw.write("<c r=\"");
//        bw.write(int2Col(column + 1));
//        bw.writeInt(rows);
//        if (StringUtil.isEmpty(s)) {
//            bw.write("\" s=\"");
//            bw.writeInt(styleIndex);
//            bw.write("\"/>");
//        } else {
//            int i;
//            if (hc.isShare() && (i = workbook.getSst().get(s)) >= 0) {
//                bw.write("\" t=\"s\" s=\"");
//                bw.writeInt(styleIndex);
//                bw.write("\"><v>");
//                bw.writeInt(i);
//                bw.write("</v></c>");
//            } else {
//                bw.write("\" t=\"inlineStr\" s=\"");
//                bw.writeInt(styleIndex);
//                bw.write("\"><is><t>");
//                bw.escapeWrite(s); // escape text
//                bw.write("</t></is></c>");
//            }
//            int ln = s.getBytes("GB2312").length; // TODO 计算
//            if (hc.width == 0 && (hc.o == null || (int) hc.o < ln)) {
//                hc.o = ln;
//            }
//        }
//    }

//    protected void writeDate(int n, int row, int column) throws IOException {
//        writeDate(date, row, column, date);
//    }

//    protected void writeDate(int n, int row, int column, int xf) throws IOException {
//        bw.write("<c r=\"");
//        bw.write(sheet.int2Col(column + 1));
//        bw.writeInt(row);
//        bw.write("\" s=\"");
//        bw.writeInt(xf);
//        bw.write("\"><v>");
//        bw.writeInt(n);
//        bw.write("</v></c>");
//    }

//    protected void writeLocalDate(LocalDate date, int column) throws IOException {
//        writeLocalDate(bw, date, column, date);
//    }

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
//
//    protected void writeTimestamp(ExtBufferedWriter bw, Timestamp ts, int column) throws IOException {
//        writeTimestamp(bw, ts, column, ts);
//    }
//
//    protected void writeTimestamp(ExtBufferedWriter bw, Timestamp ts, int column, Object o) throws IOException {
//        int styleIndex = getStyleIndex(columns[column], o);
//        bw.write("<c r=\"");
//        bw.write(int2Col(column + 1));
//        bw.writeInt(rows);
//        if (ts == null) {
//            bw.write("\" s=\"");
//            bw.writeInt(styleIndex);
//            bw.write("\"/>");
//        } else {
//            bw.write("\" s=\"");
//            bw.writeInt(styleIndex);
//            bw.write("\"><v>");
//            bw.write(toDateTimeValue(ts));
//            bw.write("</v></c>");
//        }
//    }
//
//    protected void writeLocalDateTime(ExtBufferedWriter bw, LocalDateTime ts, int column) throws IOException {
//        writeLocalDateTime(bw, ts, column, ts);
//    }
//
//    protected void writeLocalDateTime(ExtBufferedWriter bw, LocalDateTime ts, int column, Object o) throws IOException {
//        int styleIndex = getStyleIndex(columns[column], o);
//        bw.write("<c r=\"");
//        bw.write(int2Col(column + 1));
//        bw.writeInt(rows);
//        if (ts == null) {
//            bw.write("\" s=\"");
//            bw.writeInt(styleIndex);
//            bw.write("\"/>");
//        } else {
//            bw.write("\" s=\"");
//            bw.writeInt(styleIndex);
//            bw.write("\"><v>");
//            bw.write(toDateTimeValue(ts));
//            bw.write("</v></c>");
//        }
//    }
//
//    protected void writeTime(ExtBufferedWriter bw, Time date, int column) throws IOException {
//        writeTime(bw, date, column, date);
//    }
//
//    protected void writeTime(ExtBufferedWriter bw, Time date, int column, Object o) throws IOException {
//        int styleIndex = getStyleIndex(columns[column], o);
//        bw.write("<c r=\"");
//        bw.write(int2Col(column + 1));
//        bw.writeInt(rows);
//        if (date == null) {
//            bw.write("\" s=\"");
//            bw.writeInt(styleIndex);
//            bw.write("\"/>");
//        } else {
//            bw.write("\" s=\"");
//            bw.writeInt(styleIndex);
//            bw.write("\"><v>");
//            bw.write(toTimeValue(date));
//            bw.write("</v></c>");
//        }
//    }
//
//    protected void writeLocalTime(ExtBufferedWriter bw, LocalTime date, int column) throws IOException {
//        writeLocalTime(bw, date, column, date);
//    }
//
//    protected void writeLocalTime(ExtBufferedWriter bw, LocalTime date, int column, Object o) throws IOException {
//        int styleIndex = getStyleIndex(columns[column], o);
//        bw.write("<c r=\"");
//        bw.write(int2Col(column + 1));
//        bw.writeInt(rows);
//        if (date == null) {
//            bw.write("\" s=\"");
//            bw.writeInt(styleIndex);
//            bw.write("\"/>");
//        } else {
//            bw.write("\" s=\"");
//            bw.writeInt(styleIndex);
//            bw.write("\"><v>");
//            bw.write(toTimeValue(date));
//            bw.write("</v></c>");
//        }
//    }
//
//    protected void writeBigDecimal(ExtBufferedWriter bw, BigDecimal bd, int column) throws IOException {
//        writeBigDecimal(bw, bd, column, bd);
//    }
//
    private void writeDecimal(BigDecimal bd, int row, int column, int xf) throws IOException {
        bw.write("<c r=\"");
        bw.write(sheet.int2Col(column + 1));
        bw.writeInt(row);
        bw.write("\" s=\"");
        bw.writeInt(xf);
        bw.write("\"><v>");
        bw.write(bd.toString());
        bw.write("</v></c>");
    }

//    protected void writeInt(ExtBufferedWriter bw, int n, int column) throws IOException {
//        Column hc = columns[column];
//        if (hc.processor == null) {
//            writeInt0(bw, n, column);
//        } else {
//            Object o = hc.processor.conversion(n);
//            if (o != null) {
//                Class<?> clazz = o.getClass();
//                boolean blockOrDefault = blockOrDefault(hc.cellStyle);
//                if (isString(clazz)) {
//                    if (blockOrDefault) {
//                        hc.cellStyle = hc.getCellStyle(String.class);
//                    }
//                    writeString(bw, o.toString(), column, n);
//                }
//                else if (isChar(clazz)) {
//                    if (blockOrDefault) {
//                        hc.cellStyle = Styles.defaultCharBorderStyle();
//                    }
//                    char c = ((Character) o).charValue();
//                    writeChar0(bw, c, column, n);
//                }
//                else if (isInt(clazz)) {
//                    n = ((Integer) o).intValue();
//                    writeInt0(bw, n, column, n);
//                }
//                else if (isLong(clazz)) {
//                    long l = ((Long) o).longValue();
//                    writeLong(bw, l, column, n);
//                }
//                else if (isDate(clazz)) {
//                    if (blockOrDefault) {
//                        hc.cellStyle = hc.getCellStyle(java.util.Date.class);
//                    }
//                    writeDate(bw, (Date) o, column, n);
//                }
//                else if (isDateTime(clazz)) {
//                    if (blockOrDefault) {
//                        hc.cellStyle = hc.getCellStyle(java.sql.Timestamp.class);
//                    }
//                    writeTimestamp(bw, (Timestamp) o, column, n);
//                }
//                else if (isFloat(clazz)) {
//                    if (blockOrDefault) {
//                        hc.cellStyle = hc.getCellStyle(double.class);
//                    }
//                    writeDouble(bw, ((Double) o).doubleValue(), column, n);
//                }
//                else if (isBool(clazz)) {
//                    if (blockOrDefault) {
//                        hc.cellStyle = hc.getCellStyle(boolean.class);
//                    }
//                    boolean bool = ((Boolean) o).booleanValue();
//                    writeBoolean(bw, bool, column, n);
//                }
//                else if (isBigDecimal(clazz)) {
//                    if (blockOrDefault) {
//                        hc.cellStyle = hc.getCellStyle(BigDecimal.class);
//                    }
//                    writeBigDecimal(bw, (BigDecimal) o, column, n);
//                }
//                else if (isTime(clazz)) {
//                    if (blockOrDefault) {
//                        hc.cellStyle = hc.getCellStyle(Time.class);
//                    }
//                    writeTime(bw, (Time) o, column, n);
//                }
//                else  if (isLocalDate(clazz)) {
//                    if (blockOrDefault) {
//                        hc.cellStyle = hc.getCellStyle(LocalDate.class);
//                    }
//                    writeLocalDate(bw, (LocalDate) o, column, n);
//                }
//                else  if (isLocalDateTime(clazz)) {
//                    if (blockOrDefault) {
//                        hc.cellStyle = hc.getCellStyle(LocalDateTime.class);
//                    }
//                    writeLocalDateTime(bw, (LocalDateTime) o, column, n);
//                }
//                else  if (isLocalTime(clazz)) {
//                    if (blockOrDefault) {
//                        hc.cellStyle = hc.getCellStyle(LocalTime.class);
//                    }
//                    writeLocalTime(bw, (LocalTime) o, column, n);
//                }
//                else {
//                    if (blockOrDefault) {
//                        hc.cellStyle = hc.getCellStyle(String.class);
//                    }
//                    writeString(bw, o.toString(), column, n);
//                }
//            }
//            else {
//                writeNull(bw, column);
//            }
//        }
//    }
//
//    protected void writeIntAutoSize(ExtBufferedWriter bw, int n, int column) throws IOException {
//        Column hc = columns[column];
//        if (hc.processor == null) {
//            writeInt0(bw, n, column);
//        } else {
//            Object o = hc.processor.conversion(n);
//            if (o != null) {
//                Class<?> clazz = o.getClass();
//                boolean blockOrDefault = blockOrDefault(hc.cellStyle);
//                if (isString(clazz)) {
//                    if (blockOrDefault) {
//                        hc.cellStyle = hc.getCellStyle(String.class);
//                    }
//                    writeStringAutoSize(bw, o.toString(), column, n);
//                }
//                else if (isChar(clazz)) {
//                    if (blockOrDefault) {
//                        hc.cellStyle = Styles.defaultCharBorderStyle();
//                    }
//                    char c = ((Character) o).charValue();
//                    writeChar0(bw, c, column, n);
//                }
//                else if (isInt(clazz)) {
//                    int nn = ((Integer) o).intValue();
//                    writeInt0(bw, nn, column, n);
//                }
//                else if (isLong(clazz)) {
//                    long l = ((Long) o).longValue();
//                    writeLong(bw, l, column, n);
//                }
//                else if (isDate(clazz)) {
//                    if (blockOrDefault) {
//                        hc.cellStyle = hc.getCellStyle(java.util.Date.class);
//                    }
//                    writeDate(bw, (Date) o, column, n);
//                }
//                else if (isDateTime(clazz)) {
//                    if (blockOrDefault) {
//                        hc.cellStyle = hc.getCellStyle(java.sql.Timestamp.class);
//                    }
//                    writeTimestamp(bw, (Timestamp) o, column, n);
//                }
//                else if (isFloat(clazz)) {
//                    if (blockOrDefault) {
//                        hc.cellStyle = hc.getCellStyle(double.class);
//                    }
//                    writeDouble(bw, ((Double) o).doubleValue(), column, n);
//                }
//                else if (isBool(clazz)) {
//                    if (blockOrDefault) {
//                        hc.cellStyle = hc.getCellStyle(boolean.class);
//                    }
//                    boolean bool = ((Boolean) o).booleanValue();
//                    writeBoolean(bw, bool, column, n);
//                }
//                else if (isBigDecimal(clazz)) {
//                    if (blockOrDefault) {
//                        hc.cellStyle = hc.getCellStyle(BigDecimal.class);
//                    }
//                    writeBigDecimal(bw, (BigDecimal) o, column, n);
//                }
//                else if (isTime(clazz)) {
//                    if (blockOrDefault) {
//                        hc.cellStyle = hc.getCellStyle(Time.class);
//                    }
//                    writeTime(bw, (Time) o, column, n);
//                }
//                else  if (isLocalDate(clazz)) {
//                    if (blockOrDefault) {
//                        hc.cellStyle = hc.getCellStyle(LocalDate.class);
//                    }
//                    writeLocalDate(bw, (LocalDate) o, column, n);
//                }
//                else  if (isLocalDateTime(clazz)) {
//                    if (blockOrDefault) {
//                        hc.cellStyle = hc.getCellStyle(LocalDateTime.class);
//                    }
//                    writeLocalDateTime(bw, (LocalDateTime) o, column, n);
//                }
//                else  if (isLocalTime(clazz)) {
//                    if (blockOrDefault) {
//                        hc.cellStyle = hc.getCellStyle(LocalTime.class);
//                    }
//                    writeLocalTime(bw, (LocalTime) o, column, n);
//                }
//                else {
//                    if (blockOrDefault) {
//                        hc.cellStyle = hc.getCellStyle(String.class);
//                    }
//                    writeStringAutoSize(bw, o.toString(), column, n);
//                }
//            }
//            else {
//                writeNull(bw, column);
//            }
//        }
//    }
//
//    protected void writeChar(ExtBufferedWriter bw, char c, int column) throws IOException {
//        Column hc = columns[column];
//        if (hc.processor == null) {
//            writeChar0(bw, c, column);
//        } else {
//            Object o = hc.processor.conversion(c);
//            if (o != null) {
//                Class<?> clazz = o.getClass();
//                boolean blockOrDefault = blockOrDefault(hc.cellStyle);
//                if (isString(clazz)) {
//                    if (blockOrDefault) {
//                        hc.cellStyle = hc.getCellStyle(String.class);
//                    }
//                    writeString(bw, o.toString(), column, c);
//                }
//                else if (isChar(clazz)) {
//                    char cc = ((Character) o).charValue();
//                    writeChar0(bw, cc, column, c);
//                }
//                else if (isInt(clazz)) {
//                    if (blockOrDefault) {
//                        hc.cellStyle = hc.getCellStyle(int.class);
//                    }
//                    int n = ((Integer) o).intValue();
//                    writeInt0(bw, n, column, c);
//                }
//                else if (isLong(clazz)) {
//                    if (blockOrDefault) {
//                        hc.cellStyle = hc.getCellStyle(long.class);
//                    }
//                    long l = ((Long) o).longValue();
//                    writeLong(bw, l, column, c);
//                }
//                else if (isDate(clazz)) {
//                    if (blockOrDefault) {
//                        hc.cellStyle = hc.getCellStyle(java.util.Date.class);
//                    }
//                    writeDate(bw, (Date) o, column, c);
//                }
//                else if (isDateTime(clazz)) {
//                    if (blockOrDefault) {
//                        hc.cellStyle = hc.getCellStyle(java.sql.Timestamp.class);
//                    }
//                    writeTimestamp(bw, (Timestamp) o, column, c);
//                }
//                else if (isFloat(clazz)) {
//                    if (blockOrDefault) {
//                        hc.cellStyle = hc.getCellStyle(double.class);
//                    }
//                    writeDouble(bw, ((Double) o).doubleValue(), column, c);
//                }
//                else if (isBool(clazz)) {
//                    if (blockOrDefault) {
//                        hc.cellStyle = hc.getCellStyle(boolean.class);
//                    }
//                    boolean bool = ((Boolean) o).booleanValue();
//                    writeBoolean(bw, bool, column, c);
//                }
//                else if (isBigDecimal(clazz)) {
//                    if (blockOrDefault) {
//                        hc.cellStyle = hc.getCellStyle(BigDecimal.class);
//                    }
//                    writeBigDecimal(bw, (BigDecimal) o, column, c);
//                }
//                else if (isTime(clazz)) {
//                    if (blockOrDefault) {
//                        hc.cellStyle = hc.getCellStyle(Time.class);
//                    }
//                    writeTime(bw, (Time) o, column, c);
//                }
//                else  if (isLocalDate(clazz)) {
//                    if (blockOrDefault) {
//                        hc.cellStyle = hc.getCellStyle(LocalDate.class);
//                    }
//                    writeLocalDate(bw, (LocalDate) o, column, c);
//                }
//                else  if (isLocalDateTime(clazz)) {
//                    if (blockOrDefault) {
//                        hc.cellStyle = hc.getCellStyle(LocalDateTime.class);
//                    }
//                    writeLocalDateTime(bw, (LocalDateTime) o, column, c);
//                }
//                else  if (isLocalTime(clazz)) {
//                    if (blockOrDefault) {
//                        hc.cellStyle = hc.getCellStyle(LocalTime.class);
//                    }
//                    writeLocalTime(bw, (LocalTime) o, column, c);
//                }
//                else {
//                    if (blockOrDefault) {
//                        hc.cellStyle = hc.getCellStyle(String.class);
//                    }
//                    writeString(bw, o.toString(), column, c);
//                }
//            }
//            else {
//                writeNull(bw, column);
//            }
//        }
//    }
//
//    protected void writeCharAutoSize(ExtBufferedWriter bw, char c, int column) throws IOException {
//        Column hc = columns[column];
//        if (hc.processor == null) {
//            writeChar0(bw, c, column);
//        } else {
//            Object o = hc.processor.conversion(c);
//            if (o != null) {
//                Class<?> clazz = o.getClass();
//                boolean blockOrDefault = blockOrDefault(hc.cellStyle);
//                if (isString(clazz)) {
//                    if (blockOrDefault) {
//                        hc.cellStyle = hc.getCellStyle(String.class);
//                    }
//                    writeStringAutoSize(bw, o.toString(), column);
//                }
//                else if (isChar(clazz)) {
//                    char cc = ((Character) o).charValue();
//                    writeChar0(bw, cc, column, c);
//                }
//                else if (isInt(clazz)) {
//                    if (blockOrDefault) {
//                        hc.cellStyle = hc.getCellStyle(int.class);
//                    }
//                    int n = ((Integer) o).intValue();
//                    writeInt0(bw, n, column, c);
//                }
//                else if (isLong(clazz)) {
//                    if (blockOrDefault) {
//                        hc.cellStyle = hc.getCellStyle(long.class);
//                    }
//                    long l = ((Long) o).longValue();
//                    writeLong(bw, l, column, c);
//                }
//                else if (isDate(clazz)) {
//                    if (blockOrDefault) {
//                        hc.cellStyle = hc.getCellStyle(java.util.Date.class);
//                    }
//                    writeDate(bw, (Date) o, column, c);
//                }
//                else if (isDateTime(clazz)) {
//                    if (blockOrDefault) {
//                        hc.cellStyle = hc.getCellStyle(java.sql.Timestamp.class);
//                    }
//                    writeTimestamp(bw, (Timestamp) o, column, c);
//                }
//                else if (isFloat(clazz)) {
//                    if (blockOrDefault) {
//                        hc.cellStyle = hc.getCellStyle(double.class);
//                    }
//                    writeDouble(bw, ((Double) o).doubleValue(), column, c);
//                }
//                else if (isBool(clazz)) {
//                    if (blockOrDefault) {
//                        hc.cellStyle = hc.getCellStyle(boolean.class);
//                    }
//                    boolean bool = ((Boolean) o).booleanValue();
//                    writeBoolean(bw, bool, column, c);
//                }
//                else if (isBigDecimal(clazz)) {
//                    if (blockOrDefault) {
//                        hc.cellStyle = hc.getCellStyle(BigDecimal.class);
//                    }
//                    writeBigDecimal(bw, (BigDecimal) o, column, c);
//                }
//                else if (isTime(clazz)) {
//                    if (blockOrDefault) {
//                        hc.cellStyle = hc.getCellStyle(Time.class);
//                    }
//                    writeTime(bw, (Time) o, column, c);
//                }
//                else  if (isLocalDate(clazz)) {
//                    if (blockOrDefault) {
//                        hc.cellStyle = hc.getCellStyle(LocalDate.class);
//                    }
//                    writeLocalDate(bw, (LocalDate) o, column, c);
//                }
//                else  if (isLocalDateTime(clazz)) {
//                    if (blockOrDefault) {
//                        hc.cellStyle = hc.getCellStyle(LocalDateTime.class);
//                    }
//                    writeLocalDateTime(bw, (LocalDateTime) o, column, c);
//                }
//                else  if (isLocalTime(clazz)) {
//                    if (blockOrDefault) {
//                        hc.cellStyle = hc.getCellStyle(LocalTime.class);
//                    }
//                    writeLocalTime(bw, (LocalTime) o, column, c);
//                }
//                else {
//                    if (blockOrDefault) {
//                        hc.cellStyle = hc.getCellStyle(String.class);
//                    }
//                    writeStringAutoSize(bw, o.toString(), column, c);
//                }
//            } else {
//                writeNull(bw, column);
//            }
//        }
//    }
//    private void writeInt0(ExtBufferedWriter bw, int n, int column) throws IOException {
//        writeInt0(bw, n, column, n);
//    }
//
//    private void writeInt(int n, int row, int column, int xf) throws IOException {
//        bw.write("<c r=\"");
//        bw.write(sheet.int2Col(column + 1));
//        bw.writeInt(row);
//        bw.write("\" s=\"");
//        bw.writeInt(xf);
//        bw.write("\"><v>");
//        bw.writeInt(n);
//        bw.write("</v></c>");
//    }
//
//    private void writeChar0(ExtBufferedWriter bw, char c, int column) throws IOException {
//        writeChar0(bw, c, column, c);
//    }

    private void writeChar(char c, int row, int column, int xf) throws IOException {
        bw.write("<c r=\"");
        bw.write(sheet.int2Col(column + 1));
        bw.writeInt(row);
        bw.write("\" t=\"s\" s=\"");
        bw.writeInt(xf);
        bw.write("\"><v>");
        bw.writeInt(sst.get(c));
        bw.write("</v></c>");
    }

//    protected void writeLong(ExtBufferedWriter bw, long l, int column) throws IOException {
//        writeLong(bw, l, column, l);
//    }

    private void writeNumeric(long l, int row, int column, int xf) throws IOException {
        bw.write("<c r=\"");
        bw.write(sheet.int2Col(column + 1));
        bw.writeInt(row);
        bw.write("\" s=\"");
        bw.writeInt(xf);
        bw.write("\"><v>");
        bw.write(l);
        bw.write("</v></c>");
    }

//    protected void writeDouble(ExtBufferedWriter bw, double d, int column) throws IOException {
//        writeDouble(bw, d, column, d);
//    }
//
//    protected void writeDouble(ExtBufferedWriter bw, double d, int column, Object o) throws IOException {
//        int styleIndex = getStyleIndex(columns[column], o);
//        bw.write("<c r=\"");
//        bw.write(int2Col(column + 1));
//        bw.writeInt(rows);
//        bw.write("\" s=\"");
//        bw.writeInt(styleIndex);
//        bw.write("\"><v>");
//        bw.write(d);
//        bw.write("</v></c>");
//    }
//
//    protected void writeBoolean(ExtBufferedWriter bw, boolean bool, int column) throws IOException {
//        writeBoolean(bw, bool, column, bool);
//    }
//
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

    protected void writeNull(int row, int column, int xf) throws IOException {
        bw.write("<c r=\"");
        bw.write(sheet.int2Col(column + 1));
        bw.writeInt(row);
        bw.write("\" s=\"");
        bw.writeInt(xf);
        bw.write("\"/>");
    }

//    /**
//     * 写空行数据
//     * @param bw
//     */
//    protected void writeEmptyRow(ExtBufferedWriter bw) throws IOException {
//        // Row number
//        int r = ++rows;
//        final int len = columns.length;
//        bw.write("<row r=\"");
//        bw.writeInt(r);
//        bw.write("\" ht=\"16.5\" spans=\"1:");
//        bw.writeInt(len);
//        bw.write("\">");
//
//        Styles styles = workbook.getStyles();
//        for (int i = 1; i <= len; i++) {
//            Column hc = columns[i - 1];
//            bw.write("<c r=\"");
//            bw.write(int2Col(i));
//            bw.writeInt(r);
//
//            int style = hc.getCellStyle();
//            // 隔行变色
//            if (autoOdd == 0 && isOdd() && !Styles.hasFill(style)) {
//                style |= oddFill;
//            }
//            int styleIndex = styles.of(style);
//            bw.write("\" s=\"");
//            bw.writeInt(styleIndex);
//            bw.write("\"/>");
//
//            if (hc.o == null) {
//                hc.o = hc.getName().getBytes("GB2312").length;
//            }
//        }
//        bw.write("</row>");
//    }
//
//    protected  void autoColumnSize(File sheet) throws IOException {
//        // resize each column width ...
//        File temp = new File(sheet.getParent(), sheet.getName() + ".temp");
//        if (!sheet.renameTo(temp)) {
//            Files.move(sheet.toPath(), temp.toPath(), StandardCopyOption.REPLACE_EXISTING);
//        }
//
//        FileChannel inChannel = null;
//        FileChannel outChannel = null;
//        try (FileInputStream fis = new FileInputStream(temp);
//             FileOutputStream fos = new FileOutputStream(sheet)) {
//            inChannel = fis.getChannel();
//            outChannel = fos.getChannel();
//
//            inChannel.transferTo(0, headInfoLen, outChannel);
//            ByteBuffer buffer = ByteBuffer.allocate(baseInfoLen);
//            inChannel.read(buffer, headInfoLen);
//            buffer.compact();
//            byte b;
//            if ((b = buffer.get()) == '"') {
//                char[] chars = int2Col(columns.length);
//                String s = ':' + new String(chars) + rows;
//                outChannel.write(ByteBuffer.wrap(s.getBytes(Const.UTF_8)));
//            }
//            buffer.flip();
//            buffer.put(b);
//            buffer.compact();
//            outChannel.write(buffer);
//
//            StringBuilder buf = new StringBuilder();
//            buf.append("<cols>");
//            int i = 0;
//            for (Column hc : columns) {
//                i++;
//                buf.append("<col customWidth=\"1\" width=\"");
//                if (hc.width > 0.0000001) {
//                    buf.append(hc.width);
//                    buf.append("\" max=\"");
//                    buf.append(i);
//                    buf.append("\" min=\"");
//                    buf.append(i);
//                    buf.append("\"/>");
//                } else if (autoSize == 1) {
//                    int _l = hc.name.getBytes("GB2312").length, len;
//                    // TODO 根据字体字号计算文本宽度
//                    if (isString(hc.clazz)) {
//                        if (hc.o == null) {
//                            len = 0;
//                        } else {
//                            len = (int) hc.o;
//                        }
////                        len = hc.o.toString().getBytes("GB2312").length;
//                    }
//                    else if (isDate(hc.clazz) || isLocalDate(hc.clazz)) {
//                        len = 10;
//                    }
//                    else if (isDateTime(hc.clazz) || isLocalDateTime(hc.clazz)) {
//                        len = 20;
//                    }
//                    else if (isInt(hc.clazz)) {
//                        // TODO 根据numFmt计算字符宽度
//                        len = hc.type > 0 ? 12 :  11;
//                    }
//                    else if (isFloat(hc.clazz)) {
//                        // TODO 根据numFmt计算字符宽度
//                        if (hc.o == null) {
//                            len = 0;
//                        } else {
//                            len = hc.o.toString().getBytes("GB2312").length;
//                        }
////                        if (len < 11) {
////                            len = hc.type > 0 ? 12 : 11;
////                        }
//                    } else if (isTime(hc.clazz) || isLocalTime(hc.clazz)) {
//                        len = 8;
//                    } else {
//                        len = 10;
//                    }
//                    buf.append(_l > len ? _l + 3.38 : len + 3.38);
//                    buf.append("\" max=\"");
//                    buf.append(i);
//                    buf.append("\" min=\"");
//                    buf.append(i);
//                    buf.append("\" bestFit=\"1\"/>");
//                } else {
//                    buf.append(width);
//                    buf.append("\" max=\"");
//                    buf.append(i);
//                    buf.append("\" min=\"");
//                    buf.append(i);
//                    buf.append("\"/>");
//                }
//            }
//            buf.append("</cols>");
//
//            outChannel.write(ByteBuffer.wrap(buf.toString().getBytes(Const.UTF_8)));
//            int start = headInfoLen + baseInfoLen;
//            inChannel.transferTo(start, inChannel.size() - start, outChannel);
//
//        } catch (IOException e) {
//            throw e;
//        } finally {
//            boolean delete = temp.delete();
//            if (!delete) {
//                what("9005", temp.getAbsolutePath());
//            }
//            if (inChannel != null) {
//                inChannel.close();
//            }
//            if (outChannel != null) {
//                outChannel.close();
//            }
//        }
//    }

    @Override
    public void close() throws IOException {
        FileUtil.close(bw);
    }
}

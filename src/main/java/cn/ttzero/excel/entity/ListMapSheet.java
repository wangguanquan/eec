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

package cn.ttzero.excel.entity;

import cn.ttzero.excel.reader.Cell;
import cn.ttzero.excel.util.StringUtil;

import java.io.IOException;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import static cn.ttzero.excel.manager.Const.ROW_BLOCK_SIZE;

/**
 * Created by guanquan.wang at 2018-01-26 14:46
 */
public class ListMapSheet extends Sheet {
    private List<Map<String, ?>> data;

    public ListMapSheet(Workbook workbook) {
        super(workbook);
    }

    public ListMapSheet(Workbook workbook, String name, Column[] columns) {
        super(workbook, name, columns);
    }

    public ListMapSheet(Workbook workbook, String name, WaterMark waterMark, Column[] columns) {
        super(workbook, name, waterMark, columns);
    }

    /**
     * Returns the header column info
     * @return array of column
     */
    @Override
    public Column[] getHeaderColumns() {
        if (headerReady) return columns;
        @SuppressWarnings("unchecked")
        Map<String, ?> first = (Map<String, ?>) workbook.getFirst(data);
        // No data
        if (first == null) {
            if (columns == null) {
                columns = new Column[0];
            }
        }
        else if (columns.length == 0) {
            int size = first.size(), i = 0;
            columns = new Column[size];
            for (Iterator<? extends Map.Entry<String, ?>> it = first.entrySet().iterator(); it.hasNext(); ) {
                Map.Entry<String, ?> entry = it.next();
                columns[i++] = new Column(entry.getKey(), entry.getKey(), entry.getValue().getClass());
            }
        }
        else {
            for (int i = 0; i < columns.length; i++) {
                Column hc = columns[i];
                if (StringUtil.isEmpty(hc.key)) {
                    throw new ExcelWriteException(getClass() + " 类别必须指定map的key。");
                }
                if (hc.getClazz() == null) {
                    hc.setClazz(first.get(hc.key).getClass());
                }
            }
        }
        for (Column hc : columns) {
            hc.styles = workbook.getStyles();
        }
        headerReady = true;
        return columns;
    }

    @Override
    public void close() throws IOException {
        data.clear();
        data = null;
        super.close();
    }

    public ListMapSheet setData(final List<Map<String, ?>> data) {
        this.data = data;
        return this;
    }

//    @Override
//    public void writeTo(Path xl) throws IOException {
////        Path worksheets = xl.resolve("worksheets");
////        if (!Files.exists(worksheets)) {
////            FileUtil.mkdir(worksheets);
////        }
////        String name = getFileName();
////        workbook.what("0010", getName());
////
////        @SuppressWarnings("unchecked")
////        Map<String, ?> first = (Map<String, ?>) workbook.getFirst(data);
////        for (int i = 0; i < columns.length; i++) {
////            Column hc = columns[i];
////            if (StringUtil.isEmpty(hc.key)) {
////                throw new IOException(getClass() + " 类别必须指定map的key。");
////            }
////            if (hc.getClazz() == null) {
////                hc.setClazz(first.get(hc.key).getClass());
////            }
////        }
////
////        File sheetFile = worksheets.resolve(name).toFile();
////
////        // write date
////        try (ExtBufferedWriter bw = new ExtBufferedWriter(new OutputStreamWriter(new FileOutputStream(sheetFile), StandardCharsets.UTF_8))) {
////            // Write header
////            writeBefore(bw);
////            // Main data
////            // Write sheet data
////            if (getAutoSize() == 1) {
////                for (Map<String, ?> map : data) {
////                    writeRowAutoSize(map, bw);
////                }
////            } else {
////                for (Map<String, ?> map : data) {
////                    writeRow(map, bw);
////                }
////            }
////
////            // Write foot
////            writeAfter(bw);
////
////        }
////
////        // resize columns
////        boolean resize = false;
////        for (Column hc : columns) {
////            if (hc.getWidth() > 0.000001) {
////                resize = true;
////                break;
////            }
////        }
////
////        if (getAutoSize() == 1 || resize) {
////            autoColumnSize(sheetFile);
////        }
////
////        // relationship
////        relManager.write(worksheets, name);
//    }

    /**
     * Returns a row-block. The row-block is content by 32 rows
     * @return a row-block
     */
    @Override
    public RowBlock nextBlock() {
        // clear first
        rowBlock.clear();

        loopData();

        return rowBlock.flip();
    }

    private void loopData() {
        int end = getEndIndex();
        List<Map<String, ?>> sub = data.subList(rows, end);
        int len = columns.length;
        for (Map<String, ?> map : sub) {
            Row row = rowBlock.next();
            row.index = rows++;
            Cell[] cells = row.realloc(len);
            for (int i = 0; i < len; i++) {
                Column hc = columns[i];
                Object e = map.get(hc.key);
                // clear cells
                Cell cell = cells[i];
                cell.clear();

                // blank cell
                if (e == null) {
                    cell.setBlank();
                    continue;
                }

                setCellValue(cell, e, hc);
            }
        }
        if (end == data.size()) {
            rowBlock.markEnd();
        }
    }

    private int getEndIndex() {
        int end = rows + ROW_BLOCK_SIZE;
        return end <= data.size() ? end : data.size();
    }

//    /**
//     * write map
//     *
//     * @param map
//     * @param bw
//     * @throws IOException
//     */
//    protected void writeRowAutoSize(Map<String, ?> map, ExtBufferedWriter bw) throws IOException {
//        if (map == null) {
//            writeEmptyRow(bw);
//            return;
//        }
//        // 行番号
//        int r = ++rows;
//        final int len = columns.length;
//        bw.write("<row r=\"");
//        bw.writeInt(r);
//        bw.write("\" spans=\"1:");
//        bw.writeInt(len);
//        bw.write("\">");
//
//        for (int i = 0; i < len; i++) {
//            Column hc = columns[i];
//            Object o = map.get(hc.key);
//            if (o != null) {
//                Class<?> clazz = hc.getClazz();
//                // t n=numeric (default), s=string, b=boolean
//                if (isString(clazz)) {
//                    String s = o.toString();
//                    writeStringAutoSize(bw, s, i);
//                }
//                else if (isDate(clazz)) {
//                    java.util.Date date = (java.util.Date) o;
//                    writeDate(bw, date, i);
//                }
//                else if (isDateTime(clazz)) {
//                    Timestamp ts = (Timestamp) o;
//                    writeTimestamp(bw, ts, i);
//                }
//                else if (isChar(clazz)) {
//                    char c = ((Character) o).charValue();
//                    writeCharAutoSize(bw, c, i);
//                }
//                else if (isInt(clazz)) {
//                    int n = ((Integer) o).intValue();
//                    writeIntAutoSize(bw, n, i);
//                }
//                else if (isLong(clazz)) {
//                    long l = ((Long) o).longValue();
//                    writeLong(bw, l, i);
//                }
//                else if (isFloat(clazz)) {
//                    double d = ((Double) o).doubleValue();
//                    writeDouble(bw, d, i);
//                }
//                else if (isBool(clazz)) {
//                    boolean bool = ((Boolean) o).booleanValue();
//                    writeBoolean(bw, bool, i);
//                }
//                else if (isBigDecimal(clazz)) {
//                    writeBigDecimal(bw, (BigDecimal) o, i);
//                }
//                else if (isLocalDate(clazz)) {
//                    writeLocalDate(bw, (LocalDate) o, i);
//                }
//                else if (isLocalDateTime(clazz)) {
//                    writeLocalDateTime(bw, (LocalDateTime) o, i);
//                }
//                else if (isTime(clazz)) {
//                    writeTime(bw, (java.sql.Time) o, i);
//                }
//                else if (isLocalTime(clazz)) {
//                    writeLocalTime(bw, (java.time.LocalTime) o, i);
//                }
//                else {
//                    writeStringAutoSize(bw, o.toString(), i);
//                }
//            } else {
//                writeNull(bw, i);
//            }
//        }
//        bw.write("</row>");
//    }
//
//    protected void writeRow(Map<String, ?> map, ExtBufferedWriter bw) throws IOException {
//        if (map == null) {
//            writeEmptyRow(bw);
//            return;
//        }
//        // Row number
//        int r = ++rows;
//        final int len = columns.length;
//        bw.write("<row r=\"");
//        bw.writeInt(r);
//        bw.write("\" spans=\"1:");
//        bw.writeInt(len);
//        bw.write("\">");
//
//        for (int i = 0; i < len; i++) {
//            Column hc = columns[i];
//            Object o = map.get(hc.key);
//            if (o != null) {
//                Class<?> clazz = hc.getClazz();
//                // t n=numeric (default), s=string, b=boolean
//                if (isString(clazz)) {
//                    String s = o.toString();
//                    writeString(bw, s, i);
//                }
//                else if (isDate(clazz)) {
//                    java.util.Date date = (java.util.Date) o;
//                    writeDate(bw, date, i);
//                }
//                else if (isDateTime(clazz)) {
//                    Timestamp ts = (Timestamp) o;
//                    writeTimestamp(bw, ts, i);
//                }
//                else if (isChar(clazz)) {
//                    char c = ((Character) o).charValue();
//                    writeChar(bw, c, i);
//                }
//                else if (isInt(clazz)) {
//                    int n = ((Integer) o).intValue();
//                    writeInt(bw, n, i);
//                }
//                else if (isLong(clazz)) {
//                    long l = ((Long) o).longValue();
//                    writeLong(bw, l, i);
//                }
//                else if (isFloat(clazz)) {
//                    double d = ((Double) o).doubleValue();
//                    writeDouble(bw, d, i);
//                }
//                else if (isBool(clazz)) {
//                    boolean bool = ((Boolean) o).booleanValue();
//                    writeBoolean(bw, bool, i);
//                }
//                else if (isBigDecimal(clazz)) {
//                    writeBigDecimal(bw, (BigDecimal) o, i);
//                }
//                else if (isLocalDate(clazz)) {
//                    writeLocalDate(bw, (LocalDate) o, i);
//                }
//                else if (isLocalDateTime(clazz)) {
//                    writeLocalDateTime(bw, (LocalDateTime) o, i);
//                }
//                else if (isTime(clazz)) {
//                    writeTime(bw, (java.sql.Time) o, i);
//                }
//                else if (isLocalTime(clazz)) {
//                    writeLocalTime(bw, (java.time.LocalTime) o, i);
//                }
//                else {
//                    writeString(bw, o.toString(), i);
//                }
//            }
//            else {
//                writeNull(bw, i);
//            }
//        }
//        bw.write("</row>");
//    }

    /**
     * Returns total rows in this worksheet
     * @return -1 if unknown
     */
    public int size() {
        return data != null ? data.size() : 0;
    }
}

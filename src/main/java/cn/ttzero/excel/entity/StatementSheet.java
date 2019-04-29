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

import cn.ttzero.excel.manager.Const;
import cn.ttzero.excel.reader.Cell;
import cn.ttzero.excel.util.ExtBufferedWriter;
import cn.ttzero.excel.util.FileUtil;
import cn.ttzero.excel.util.StringUtil;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.util.function.Supplier;

import static cn.ttzero.excel.manager.Const.ROW_BLOCK_SIZE;

/**
 * Created by guanquan.wang on 2017/9/26.
 */
public class StatementSheet extends Sheet {
    protected PreparedStatement ps;
    private ResultSet rs;

    public StatementSheet(Workbook workbook, String name, Column[] columns) {
        super(workbook, name, columns);
    }

    public StatementSheet(Workbook workbook, String name, WaterMark waterMark, Column[] columns) {
        super(workbook, name, waterMark, columns);
    }

    /**
     * @param ps PreparedStatement
     */
    public void setPs(PreparedStatement ps) {
        this.ps = ps;
    }

    /**
     * 关闭外部源
     */
    @Override
    public void close() throws IOException {
//        super.close();
        if (ps != null) {
            try {
                ps.close();
            } catch (SQLException e) {
                workbook.what("9006", e.getMessage());
            }
        }
        super.close();
    }


//    /**
//     * 写数据
//     * @param xl the storage path
//     * @throws IOException
//     * @throws ExcelWriteException
//     */
//    @Override
//    public void writeTo(Path xl) throws IOException, ExcelWriteException {
//
////        Path worksheets = xl.resolve("worksheets");
////        if (!Files.exists(worksheets)) {
////            FileUtil.mkdir(worksheets);
////        }
////        String name = getFileName();
////        workbook.what("0010", getName());
////
////        // TODO 1.判断各sheet抽出的数据量大小
////        // TODO 2.如果量大则抽取类型为String的列判断重复率
////
//        int i = 0;
//        try {
//            ResultSetMetaData metaData = ps.getMetaData();
//            for ( ; i < columns.length; i++) {
//                if (StringUtil.isEmpty(columns[i].getName())) {
//                    columns[i].setName(metaData.getColumnName(i));
//                }
//                // TODO metaData.getColumnType()
//            }
//        } catch (SQLException e) {
//        }
//
//        for (i = 0 ; i < columns.length; i++) {
//            if (StringUtil.isEmpty(columns[i].getName())) {
//                columns[i].setName(String.valueOf(i));
//            }
//        }
//
//        ResultSet rs = null;
//        try {
//            workbook.what("0011");
//            rs = ps.executeQuery();
//            if (rs.next()) {
//                RowBlock rowBlock = new RowBlock();
//                sheetWriter.write(xl, () -> {
//                    rowBlock.clear();
//
//                    // TODO loop
//                    if (sheetWriter.outOfSheet(rowBlock.getTotal())) {
//                        rowBlock.markEnd();
//                        // TODO break
//                    }
//
//                    return rowBlock;
//                });
//                // TODO paging
//            } else writeEmptySheet(xl);
//
//        } catch (SQLException e) {
//            if (rs != null) {
//                try {
//                    rs.close();
//                } catch (SQLException ex) {
//                    workbook.what("9006", ex.getMessage());
//                }
//            }
//            close();
//            throw new ExcelWriteException(e);
//        } finally {
//            sheetWriter.close();
//        }
//
////        File sheetFile = worksheets.resolve(name).toFile();
////        ResultSet rs = null;
////        int sub = 0;
////        // write date
////        try (ExtBufferedWriter bw = new ExtBufferedWriter(new OutputStreamWriter(new FileOutputStream(sheetFile), StandardCharsets.UTF_8))) {
////            workbook.what("0011");
////            rs = ps.executeQuery();
////            workbook.what("0012");
////            // Write header
////            writeBefore(bw);
////            int limit = Const.Limit.MAX_ROWS_ON_SHEET - rows; // exclude header rows
////            if (rs.next()) {
////                // Write sheet data
////                if (getAutoSize() == 1) {
////                    do {
////                        // Paging
////                        if (rows >= limit) {
////                            writeRowAutoSize(rs, bw);
////                            sub++;
////                            break;
////                        }
////                        writeRowAutoSize(rs, bw);
////                    } while (rs.next());
////                } else {
////                    do {
////                        // Paging
////                        if (rows >= limit) {
////                            writeRow(rs, bw);
////                            sub++;
////                            break;
////                        }
////                        writeRow(rs, bw);
////                    } while (rs.next());
////                }
////            }
////
////            // Write foot
////            writeAfter(bw);
////
////        } catch (SQLException e) {
////            close();
////            if (rs != null) {
////                try {
////                    rs.close();
////                } catch (SQLException ex) {
////                    workbook.what("9006", ex.getMessage());
////                }
////            }
////            throw new ExcelWriteException(e);
////        } finally {
////            if (rows < Const.Limit.MAX_ROWS_ON_SHEET) {
////                if (rs != null) {
////                    try {
////                        rs.close();
////                    } catch (SQLException e) {
////                        workbook.what("9006", e.getMessage());
////                    }
////                }
////                close();
////            }
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
////        boolean autoSize;
////        if (autoSize = (getAutoSize() == 1 || resize)) {
////            autoColumnSize(sheetFile);
////        }
////
////        // relationship
////        relManager.write(worksheets, name);
////
////        // Paging
////        if (sub == 1) {
////            workbook.what("0013");
////            ResultSetSheet rss = new ResultSetSheet(workbook, this.name + " (" + (sub) + ")", waterMark, columns, rs, relManager.clone());
////            rss.setCopySheet(true);
////            if (autoSize) rss.autoSize();
////            rss.autoOdd = this.autoOdd;
////            rss.oddFill = this.oddFill;
////            workbook.insertSheet(id, rss);
////            rss.writeTo(xl);
////        }
////
////        close();
//
//    }

    /**
     * write worksheet data to path
     * @param path the storage path
     * @throws IOException write error
     * @throws ExcelWriteException others
     */
    public void writeTo(Path path) throws IOException, ExcelWriteException {
        if (sheetWriter != null) {
            try {
                rs = ps.executeQuery();
            } catch (SQLException e) {
                throw new ExcelWriteException(e);
            }
            rowBlock = new RowBlock();
            sheetWriter.write(path);
        } else {
            throw new ExcelWriteException("Worksheet writer is not instanced.");
        }
    }

    /**
     * Returns a row-block. The row-block is content by 32 rows
     * @return a row-block
     */
    @Override
    public RowBlock nextBlock() {
        // clear first
        rowBlock.clear();

        try {
            loopData();
        } catch (SQLException e) {
            throw new ExcelWriteException(e);
        }

        // TODO paging

        return rowBlock.flip();
    }

    private void loopData() throws SQLException {
        int len = columns.length, n = 0;

        for (; n++ < ROW_BLOCK_SIZE && rs.next(); ) {
            Row row = rowBlock.next();
            row.index = rows++;
            Cell[] cells = row.realloc(len);
            for (int i = 1; i <= len; i++) {
                Column hc = columns[i - 1];

                // clear cells
                Cell cell = cells[i - 1];
                cell.clear();

                Object e = rs.getObject(i);

//                Class<?> clazz = hc.clazz;
//
//                if (isString(clazz)) {
//                    cell.setSv(rs.getString(i));
//                } else if (isDate(clazz)) {
//                    cell.setAv(toDateValue(rs.getDate(i)));
//                } else if (isDateTime(clazz)) {
//                    cell.setIv(toDateTimeValue(rs.getTimestamp(i)));
//                } else if (isChar(clazz)) {
//                    String s = rs.getString(i);
//                    cell.setCv(();
//                } else if (isShort(clazz)) {
//                    cell.setNv((Short) e);
//                } else if (isInt(clazz)) {
//                    cell.setNv((Integer) e);
//                } else if (isLong(clazz)) {
//                    cell.setLv((Long) e);
//                } else if (isFloat(clazz)) {
//                    cell.setDv((Float) e);
//                } else if (isDouble(clazz)) {
//                    cell.setDv((Double) e);
//                } else if (isBool(clazz)) {
//                    cell.setBv((Boolean) e);
//                } else if (isBigDecimal(clazz)) {
//                    cell.setMv((BigDecimal) e);
//                } else if (isLocalDate(clazz)) {
//                    cell.setAv(toDateValue((java.time.LocalDate) e));
//                } else if (isLocalDateTime(clazz)) {
//                    cell.setIv(toDateTimeValue((java.time.LocalDateTime) e));
//                } else if (isTime(clazz)) {
//                    cell.setTv(toTimeValue((java.sql.Time) e));
//                } else if (isLocalTime(clazz)) {
//                    cell.setTv(toTimeValue((java.time.LocalTime) e));
//                } else {
//                    cell.setSv(e.toString());
//                }
//                cell.xf = getStyleIndex(hc, e);

                // blank cell
                if (e == null) {
                    cell.setBlank();
                    continue;
                }

                setCellValue(cell, e, hc);
            }
        }
    }


    @Override
    public Column[] getHeaderColumns() {
        if (headerReady) return columns;
        int i = 0;
        try {
            ResultSetMetaData metaData = ps.getMetaData();
            for ( ; i < columns.length; i++) {
                if (StringUtil.isEmpty(columns[i].getName())) {
                    columns[i].setName(metaData.getColumnName(i));
                }
                // TODO metaData.getColumnType()
            }
        } catch (SQLException e) {
        }

        for (i = 0 ; i < columns.length; i++) {
            if (StringUtil.isEmpty(columns[i].getName())) {
                columns[i].setName(String.valueOf(i));
            }
        }
        return columns;
    }
}

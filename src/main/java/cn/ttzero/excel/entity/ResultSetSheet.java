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

import cn.ttzero.excel.manager.RelManager;
import cn.ttzero.excel.reader.Cell;
import cn.ttzero.excel.util.StringUtil;

import java.io.IOException;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Timestamp;
import java.util.List;
import java.util.Map;

import static cn.ttzero.excel.entity.IWorksheetWriter.*;
import static cn.ttzero.excel.entity.IWorksheetWriter.isLocalTime;
import static cn.ttzero.excel.entity.IWorksheetWriter.isTime;
import static cn.ttzero.excel.manager.Const.ROW_BLOCK_SIZE;
import static cn.ttzero.excel.util.DateUtil.toDateTimeValue;
import static cn.ttzero.excel.util.DateUtil.toDateValue;
import static cn.ttzero.excel.util.DateUtil.toTimeValue;

/**
 * Created by guanquan.wang on 2017/9/27.
 */
public class ResultSetSheet extends Sheet {
    private ResultSet rs;
    private boolean copySheet;

    public ResultSetSheet(Workbook workbook) {
        super(workbook);
    }

    public ResultSetSheet(Workbook workbook, String name, Column[] columns) {
        super(workbook, name, columns);
    }

    public ResultSetSheet(Workbook workbook, String name, WaterMark waterMark, Column[] columns) {
        super(workbook, name, waterMark, columns);
    }

    public ResultSetSheet(Workbook workbook, String name, WaterMark waterMark, Column[] columns, ResultSet rs, RelManager relManager) {
        super(workbook, name, waterMark, columns);
        this.rs = rs;
        this.relManager = relManager.clone();
    }

    public void setRs(ResultSet rs) {
        this.rs = rs;
    }

    public ResultSetSheet setCopySheet(boolean copySheet) {
        this.copySheet = copySheet;
        return this;
    }

    @Override
    public void close() throws IOException {
//        super.close();
        if (rs != null) {
            try {
                rs.close();
            } catch (SQLException e) {
                workbook.what("9006", e.getMessage());
            }
        }
        super.close();
    }

//    @Override
//    public void writeTo(Path xl) throws IOException, ExcelWriteException {
//        Path worksheets = xl.resolve("worksheets");
//        if (!Files.exists(worksheets)) {
//            FileUtil.mkdir(worksheets);
//        }
//        String name = getFileName();
//        workbook.what("0010", getName());
//
//        for (int i = 0; i < columns.length; i++) {
//            if (StringUtil.isEmpty(columns[i].getName())) {
//                columns[i].setName(String.valueOf(i));
//            }
//        }
//
//        boolean paging = false;
////        File sheetFile = worksheets.resolve(name).toFile();
////        // write date
////        try (ExtBufferedWriter bw = new ExtBufferedWriter(new OutputStreamWriter(new FileOutputStream(sheetFile), StandardCharsets.UTF_8))) {
////            // Write header
////            writeBefore(bw);
////            int limit = Const.Limit.MAX_ROWS_ON_SHEET - rows; // exclude header rows
////            // Main data
////            if (rs != null && rs.next()) {
////
////                // Write sheet data
////                if (getAutoSize() == 1) {
////                    do {
////                        // row >= max rows
////                        if (rows >= limit) {
////                            paging = !paging;
////                            writeRowAutoSize(rs, bw);
////                            break;
////                        }
////                        writeRowAutoSize(rs, bw);
////                    } while (rs.next());
////                } else {
////                    do {
////                        // Paging
////                        if (rows >= limit) {
////                            paging = !paging;
////                            writeRow(rs, bw);
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
////            throw new ExcelWriteException(e);
////        } finally {
////            if (rows < Const.Limit.MAX_ROWS_ON_SHEET) {
////                close();
////            }
////        }
////
////        // Delete empty copy sheet
////        if (copySheet && rows == 1) {
////            workbook.remove(id - 1);
////            sheetFile.delete();
////            return;
////        }
////
////        // resize columns
////        boolean resize = false;
////        for  (Column hc : columns) {
////            if (hc.getWidth() > 0.000001) {
////                resize = true;
////                break;
////            }
////        }
////        if (getAutoSize() == 1 || resize) {
////            autoColumnSize(sheetFile);
////        }
////
////        // relationship
////        relManager.write(worksheets, name);
////
////        if (paging) {
////            workbook.what("0013");
////            int sub;
////            if (!copySheet) {
////                sub = 0;
////            } else {
////                sub = Integer.parseInt(this.name.substring(this.name.lastIndexOf('(') + 1, this.name.lastIndexOf(')')));
////            }
////            String sheetName = this.name;
////            if (copySheet) {
////                sheetName = sheetName.substring(0, sheetName.lastIndexOf('(') - 1);
////            }
////
////            ResultSetSheet rss = clone();
////            rss.name = sheetName + " (" + (sub + 1) + ")";
////            workbook.insertSheet(id, rss);
////            rss.writeTo(xl);
////        }
//
//    }

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

    /**
     * Returns the header column info
     * @return array of column
     */
    @Override
    public Column[] getHeaderColumns() {
        if (headerReady) return columns;
        if (columns != null) {
            for (int i = 0; i < columns.length; i++) {
                if (StringUtil.isEmpty(columns[i].getName())) {
                    columns[i].setName(String.valueOf(i));
                }
            }
        } else columns = new Column[0];
        return columns;
    }

    protected ResultSetSheet copy() {
        ResultSetSheet rss =  new ResultSetSheet(workbook, name, waterMark, columns, rs, relManager).setCopySheet(true);
        if (getAutoSize() == 1) {
            rss.autoSize();
        }
        rss.autoOdd = this.autoOdd;
        rss.oddFill = this.oddFill;
        return rss;
    }
}

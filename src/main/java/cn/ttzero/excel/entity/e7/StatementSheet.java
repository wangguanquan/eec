/*
 * Copyright (c) 2009, guanquan.wang@yandex.com All Rights Reserved.
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

import cn.ttzero.excel.manager.Const;
import cn.ttzero.excel.entity.ExportException;
import cn.ttzero.excel.entity.WaterMark;
import cn.ttzero.excel.util.ExtBufferedWriter;
import cn.ttzero.excel.util.FileUtil;
import cn.ttzero.excel.util.StringUtil;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;

/**
 * Created by guanquan.wang at 2017/9/26.
 */
public class StatementSheet extends Sheet {
    private PreparedStatement ps;

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
    public void close() {
//        super.close();
        if (ps != null) {
            try {
                ps.close();
            } catch (SQLException e) {
                workbook.what("9006", e.getMessage());
            }
        }
    }

    /**
     * 写数据
     * @param xl sheet.xml path
     * @throws IOException
     * @throws ExportException
     */
    @Override
    public void writeTo(Path xl) throws IOException, ExportException {
        Path worksheets = xl.resolve("worksheets");
        if (!Files.exists(worksheets)) {
            FileUtil.mkdir(worksheets);
        }
        String name = getFileName();
        workbook.what("0010", getName());

        // TODO 1.判断各sheet抽出的数据量大小
        // TODO 2.如果量大则抽取类型为String的列判断重复率

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

        File sheetFile = worksheets.resolve(name).toFile();
        ResultSet rs = null;
        int sub = 0;
        // write date
        try (ExtBufferedWriter bw = new ExtBufferedWriter(new OutputStreamWriter(new FileOutputStream(sheetFile), StandardCharsets.UTF_8))) {
            workbook.what("0011");
            rs = ps.executeQuery();
            workbook.what("0012");
            // Write header
            writeBefore(bw);
            int limit = Const.Limit.MAX_ROWS_ON_SHEET_07 - rows; // exclude header rows
            if (rs.next()) {
                // Write sheet data
                if (getAutoSize() == 1) {
                    do {
                        // Paging
                        if (rows >= limit) {
                            writeRowAutoSize(rs, bw);
                            sub++;
                            break;
                        }
                        writeRowAutoSize(rs, bw);
                    } while (rs.next());
                } else {
                    do {
                        // Paging
                        if (rows >= limit) {
                            writeRow(rs, bw);
                            sub++;
                            break;
                        }
                        writeRow(rs, bw);
                    } while (rs.next());
                }
            }

            // Write foot
            writeAfter(bw);

        } catch (SQLException e) {
            close();
            if (rs != null) {
                try {
                    rs.close();
                } catch (SQLException ex) {
                    workbook.what("9006", ex.getMessage());
                }
            }
            throw new ExportException(e);
        } finally {
            if (rows < Const.Limit.MAX_ROWS_ON_SHEET_07) {
                if (rs != null) {
                    try {
                        rs.close();
                    } catch (SQLException e) {
                        workbook.what("9006", e.getMessage());
                    }
                }
                close();
            }
        }

        // resize columns
        boolean resize = false;
        for (Column hc : columns) {
            if (hc.getWidth() > 0.000001) {
                resize = true;
                break;
            }
        }
        boolean autoSize;
        if (autoSize = (getAutoSize() == 1 || resize)) {
            autoColumnSize(sheetFile);
        }

        // relationship
        relManager.write(worksheets, name);

        // Paging
        if (sub == 1) {
            workbook.what("0013");
            ResultSetSheet rss = new ResultSetSheet(workbook, this.name + " (" + (sub) + ")", waterMark, columns, rs, relManager.clone());
            rss.setCopySheet(true);
            if (autoSize) rss.autoSize();
            rss.autoOdd = this.autoOdd;
            rss.oddFill = this.oddFill;
            workbook.insertSheet(id, rss);
            rss.writeTo(xl);
        }

        close();
    }

}

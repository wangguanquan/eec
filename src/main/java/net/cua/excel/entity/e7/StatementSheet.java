package net.cua.excel.entity.e7;

import net.cua.excel.manager.Const;
import net.cua.excel.entity.ExportException;
import net.cua.excel.entity.WaterMark;
import net.cua.excel.util.ExtBufferedWriter;
import net.cua.excel.util.StringUtil;

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

    public void setPs(PreparedStatement ps) {
        this.ps = ps;
    }

    public PreparedStatement getPs() {
        return ps;
    }

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

    @Override
    public void writeTo(Path xl) throws IOException, ExportException {
        Path worksheets = xl.resolve("worksheets");
        if (!Files.exists(worksheets)) {
            Files.createDirectory(worksheets);
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
            columns[i].setName(String.valueOf(i));
        }

        File sheetFile = worksheets.resolve(name).toFile();
        ResultSet rs = null;
        int sub = 0;
        // write date
        try (ExtBufferedWriter bw = new ExtBufferedWriter(new OutputStreamWriter(new FileOutputStream(sheetFile), StandardCharsets.UTF_8))) {
            rs = ps.executeQuery();
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

        if (sub == 1) {
            ResultSetSheet rss = new ResultSetSheet(workbook, this.name, waterMark, columns, rs, relManager.clone());
            rss.setName(this.name + " (" + (sub) + ")");
            rss.setCopySheet(true);
            if (autoSize) rss.autoSize();
            workbook.insertSheet(id, rss);
            rss.writeTo(xl);
        }

        close();
    }

}

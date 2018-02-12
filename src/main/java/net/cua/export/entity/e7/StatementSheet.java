package net.cua.export.entity.e7;

import net.cua.export.entity.WaterMark;
import net.cua.export.manager.Const;
import net.cua.export.util.ExtBufferedWriter;
import net.cua.export.util.StringUtil;
import org.apache.log4j.Logger;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;

/**
 * Created by guanquan.wang on 2017/9/26.
 */
public class StatementSheet extends Sheet {
    private PreparedStatement ps;

    public StatementSheet(Workbook workbook, String name, HeadColumn[] headColumns) {
        super(workbook, name, headColumns);
    }

    public StatementSheet(Workbook workbook, String name, WaterMark waterMark, HeadColumn[] headColumns) {
        super(workbook, name, waterMark, headColumns);
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
                e.printStackTrace();
            }
        }
    }

    @Override
    public void writeTo(Path xl) throws IOException {
        Path worksheets = Paths.get(xl.toString(), "worksheets");
        if (!Files.exists(worksheets)) {
            Files.createDirectory(worksheets);
        }
        String name = getFileName();
//        logger.info(getName() + " | " + name);

        // TODO 1.判断各sheet抽出的数据量大小
        // TODO 2.如果量大则抽取类型为String的列判断重复率

        int i = 0;
        try {
            ResultSetMetaData metaData = ps.getMetaData();
            for ( ; i < headColumns.length; i++) {
                if (StringUtil.isEmpty(headColumns[i].getName())) {
                    headColumns[i].setName(metaData.getColumnName(i));
                }
            }
        } catch (SQLException e) {
            headColumns[i].setName(String.valueOf(i));
        }

        File sheetFile = Paths.get(worksheets.toString(), name).toFile();
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

        } catch (IOException e) {
            throw e;
        } catch (SQLException e) {
            logger.error(e);
        } finally {
            if (rows < Const.Limit.MAX_ROWS_ON_SHEET_07) {
                try {
                    rs.close();
                } catch (SQLException e) {
                    e.printStackTrace();
                }
                close();
            }
        }

        // resize columns
        boolean resize = false;
        for  (HeadColumn hc : headColumns) {
            if (hc.getWidth() > 0.000001) {
                resize = true;
                break;
            }
        }
        if (getAutoSize() == 1 || resize) {
            autoColumnSize(sheetFile);
        }

        // relationship
        relManager.write(worksheets, name);

        if (sub == 1) {
            ResultSetSheet rss = new ResultSetSheet(workbook, this.name, waterMark, headColumns, rs, relManager.clone());
            rss.setName(this.name + " (" + (sub) + ")");
            rss.setCopySheet(true);
            Sheet subSheet = workbook.insertSheet(id, rss);
            subSheet.writeTo(xl);
        }

        close();
    }

}

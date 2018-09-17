package net.cua.export.entity.e7;

import net.cua.export.entity.ExportException;
import net.cua.export.entity.WaterMark;
import net.cua.export.manager.Const;
import net.cua.export.manager.RelManager;
import net.cua.export.util.ExtBufferedWriter;
import net.cua.export.util.StringUtil;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.sql.ResultSet;
import java.sql.SQLException;

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
    public void close() {
//        super.close();
        if (rs != null) {
            try {
                rs.close();
            } catch (SQLException e) {
                logger.error(e.getErrorCode(), e);
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
        logger.debug(getName() + " | " + name);

        for (int i = 0; i < columns.length; i++) {
            if (StringUtil.isEmpty(columns[i].getName())) {
                columns[i].setName(String.valueOf(i));
            }
        }

        boolean paging = false;
        File sheetFile = worksheets.resolve(name).toFile();
        // write date
        try (ExtBufferedWriter bw = new ExtBufferedWriter(new OutputStreamWriter(new FileOutputStream(sheetFile), StandardCharsets.UTF_8))) {
            // Write header
            writeBefore(bw);
            int limit = Const.Limit.MAX_ROWS_ON_SHEET_07 - rows; // exclude header rows
            // Main data
            if (rs != null && rs.next()) {

                // Write sheet data
                if (getAutoSize() == 1) {
                    do {
                        // row >= max rows
                        if (rows >= limit) {
                            paging = !paging;
                            writeRowAutoSize(rs, bw);
                            break;
                        }
                        writeRowAutoSize(rs, bw);
                    } while (rs.next());
                } else {
                    do {
                        // Paging
                        if (rows >= limit) {
                            paging = !paging;
                            writeRow(rs, bw);
                            break;
                        }
                        writeRow(rs, bw);
                    } while (rs.next());
                }
            }

            // Write foot
            writeAfter(bw);

        } catch (SQLException e) {
            throw new ExportException(e);
        } finally {
            if (rows < Const.Limit.MAX_ROWS_ON_SHEET_07) {
                close();
            }
        }

        // Delete empty copy sheet
        if (copySheet && rows == 1) {
            logger.debug("Delete empty copy sheet");
            workbook.remove(id - 1);
            sheetFile.delete();
            return;
        }

        // resize columns
        boolean resize = false;
        for  (Column hc : columns) {
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

        if (paging) {
            int sub;
            if (!copySheet) {
                sub = 0;
            } else {
                sub = Integer.parseInt(this.name.substring(this.name.lastIndexOf('(') + 1, this.name.lastIndexOf(')')));
            }
            String sheetName = this.name;
            if (copySheet) {
                sheetName = sheetName.substring(0, sheetName.lastIndexOf('(') - 1);
            }

            ResultSetSheet rss = clone();
            rss.name = sheetName + " (" + (sub + 1) + ")";
            workbook.insertSheet(id, rss);
            rss.writeTo(xl);
        }

    }

    public ResultSetSheet clone() {
        ResultSetSheet rss =  new ResultSetSheet(workbook, name, waterMark, columns, rs, relManager).setCopySheet(true);
        if (getAutoSize() == 1) {
            rss.autoSize();
        }
        return rss;
    }
}

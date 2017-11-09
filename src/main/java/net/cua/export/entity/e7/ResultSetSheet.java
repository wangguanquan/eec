package net.cua.export.entity.e7;

import net.cua.export.manager.Const;
import net.cua.export.util.ExtBufferedWriter;
import org.apache.log4j.Logger;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.nio.charset.StandardCharsets;
import java.sql.ResultSet;
import java.sql.SQLException;

/**
 * Created by wanggq on 2017/9/27.
 */
public class ResultSetSheet extends Sheet {
    private Logger logger = Logger.getLogger(this.getClass().getName());
    private ResultSet rs;
    private boolean copySheet;

    public ResultSetSheet(Workbook workbook) {
        super(workbook);
    }

    public ResultSetSheet(Workbook workbook, String name, HeadColumn[] headColumns) {
        super(workbook, name, headColumns);
    }

    public ResultSetSheet(Workbook workbook, String name, String waterMark, HeadColumn[] headColumns) {
        super(workbook, name, waterMark, headColumns);
    }

    public void setRs(ResultSet rs) {
        this.rs = rs;
    }

    public void setCopySheet(boolean copySheet) {
        this.copySheet = copySheet;
    }

    @Override
    public void close() {
        super.close();
        if (rs != null) {
            try {
                rs.close();
            } catch (SQLException e) {
                logger.error(e.getErrorCode(), e);
            }
        }
    }

    @Override
    public void writeTo(File root) {
        logger.info(getName());

        File xl = new File(root, "xl");
        if (!xl.exists() && !xl.mkdirs()) {
            // TODO echo error
            return;
        }

        File parent = new File(xl, "worksheets");
        if (!parent.exists() && !parent.mkdir()) {
            // TODO echo error
            return;
        }
        String name = getFileName();

        // relationship
        try {
            relManager.write(parent, name);
        } catch (IllegalAccessException e) {
            e.printStackTrace();
        }

        File sheetFile = new File(parent, name);

        // write date
        try (ExtBufferedWriter bw = new ExtBufferedWriter(new OutputStreamWriter(new FileOutputStream(sheetFile), StandardCharsets.UTF_8));) {
            // Write header
            writeBefore(bw);
            // Main data
            if (rs != null && rs.next()) {
                // Shared string
                SharedStrings sst = workbook.getSst();
                Styles styles = workbook.getStyles();
                int sub;
                if (!copySheet) {
                    sub = 0;
                } else {
                    sub = Integer.parseInt(this.name.substring(this.name.lastIndexOf('(') + 1, this.name.lastIndexOf(')')));
//                    if (sub > 1) {
//                        sub--;
//                    }
                }
                // Write sheet data
                if (getAutoSize() == 1) {
                    do {
                        // TODO row > max rows
                        if (rows > Const.Limit.MAX_ROWS_ON_SHEET) {
                            // TODO insert sub sheet
                            ResultSetSheet rss = new ResultSetSheet(workbook, this.name, waterMark, headColumns);
                            String sheetName = this.name;
                            if (copySheet) {
                                sheetName = sheetName.substring(0, sheetName.lastIndexOf('(') - 1);
                            }
                            rss.setName(sheetName + " (" + (sub + 1) + ")");
                            rss.setRs(rs);
                            rss.setCopySheet(true);
                            if (sub > 0) {
                                sub--;
                            }
                            Sheet subSheet = workbook.insertSheet(id, rss);
                            subSheet.writeTo(root);
                            break;
                        }
                        writeRowAutoSize(rs, bw, sst, styles);
                    } while (rs.next());
                } else {
                    do {
                        // Paging
                        if (rows > Const.Limit.MAX_ROWS_ON_SHEET) {
                            // TODO insert sub sheet
                            ResultSetSheet rss = new ResultSetSheet(workbook, this.name, waterMark, headColumns);
                            String sheetName = this.name;
                            if (copySheet) {
                                sheetName = sheetName.substring(0, sheetName.lastIndexOf('(') - 1);
                            }
                            rss.setName(sheetName + " (" + (sub + 1) + ")");
                            rss.setRs(rs);
                            rss.setCopySheet(true);
                            if (sub > 0) {
                                sub--;
                            }
                            Sheet subSheet = workbook.insertSheet(id, rss);
                            subSheet.writeTo(root);
                            break;
                        }
                        writeRow(rs, bw, sst, styles);
                    } while (rs.next());
                }

//                if (sub > 1) {
//                    this.name += "(1)";
//                }
//                if (getAutoSize() == 1) {
//                    do {
//                        writeRowAutoSize(rs, bw, sst, styles);
//                    } while (rs.next());
//                } else {
//                    do {
//                        writeRow(rs, bw, sst, styles);
//                    } while (rs.next());
//                }
            }

            // Write foot
            writeAfter(bw);
        } catch (IOException e) {
            e.printStackTrace();
        } catch (SQLException e) {
            e.printStackTrace();
        } finally {
//            close();
            // TODO close ResultSet
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
    }

}

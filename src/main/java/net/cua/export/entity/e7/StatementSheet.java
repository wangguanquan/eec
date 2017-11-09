package net.cua.export.entity.e7;

import net.cua.export.manager.Const;
import net.cua.export.util.ExtBufferedWriter;
import org.apache.log4j.Logger;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;

/**
 * Created by wanggq on 2017/9/26.
 */
public class StatementSheet extends Sheet {
    private Logger logger = Logger.getLogger(this.getClass().getName());

    private PreparedStatement ps;

    public StatementSheet(Workbook workbook, String name, HeadColumn[] headColumns) {
        super(workbook, name, headColumns);
    }

    public StatementSheet(Workbook workbook, String name, String waterMark, HeadColumn[] headColumns) {
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
        super.close();
        if (ps != null) {
            try {
                ps.close();
            } catch (SQLException e) {
                e.printStackTrace();
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

        // TODO 1.判断各sheet抽出的数据量大小
        // TODO 2.如果量大则抽取类型为String的列判断重复率

        // relationship
        try {
            relManager.write(parent, name);
        } catch (IllegalAccessException e) {
            e.printStackTrace();
        }

        File sheetFile = new File(parent, name);
        ResultSet rs = null;
        int sub = 0;
        // write date
        try (ExtBufferedWriter bw = new ExtBufferedWriter(new OutputStreamWriter(new FileOutputStream(sheetFile), StandardCharsets.UTF_8))) {
            rs = ps.executeQuery();
            // Write header
            writeBefore(bw);
            if (rs.next()) {
                // Shared string
                SharedStrings sst = workbook.getSst();
                Styles styles = workbook.getStyles();
                // Write sheet data
                if (getAutoSize() == 1) {
                    do {
                        // TODO row > max rows
                        // 这里会丢数据
                        if (rows >= Const.Limit.MAX_ROWS_ON_SHEET) {
                            // TODO insert sub sheet

                            sub++;
                            break;
                        }
                        writeRowAutoSize(rs, bw, sst, styles);
                    } while (rs.next());
                } else {
                    do {
                        // Paging
                        if (rows >= Const.Limit.MAX_ROWS_ON_SHEET) {
                            // TODO insert sub sheet

                            sub++;
                            break;
                        }
                        writeRow(rs, bw, sst, styles);
                    } while (rs.next());
                }

//                // Rename self sheet
//                if (sub == 1) {
//                    this.name += "(1)";
//                }
            }

            // Write foot
            writeAfter(bw);
        } catch (IOException e) {
            e.printStackTrace();
        } catch (SQLException e) {
            e.printStackTrace();
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

        if (sub == 1) {
            ResultSetSheet rss = new ResultSetSheet(workbook, this.name, waterMark, headColumns);
            rss.setName(this.name + " (" + (sub + 1) + ")");
            rss.setRs(rs);
            rss.setCopySheet(true);
            Sheet subSheet = workbook.insertSheet(id, rss);
            subSheet.writeTo(root);
        }
    }

}

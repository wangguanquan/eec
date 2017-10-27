package net.cua.export;

import net.cua.export.entity.TooManyColumnsException;
import net.cua.export.entity.e7.SharedStrings;
import net.cua.export.entity.e7.Styles;
import net.cua.export.manager.Const;
import net.cua.export.entity.e7.Relationship;
import net.cua.export.entity.e7.Sheet;
import net.cua.export.entity.e7.Workbook;
import net.cua.export.manager.WorkbookManager;
import net.cua.export.processor.ParamProcessor;
import net.cua.export.util.FileUtil;
import net.cua.export.util.StringUtil;
import net.cua.export.util.ZipCompressor;
import org.apache.log4j.Logger;
import org.springframework.stereotype.Component;

import java.io.File;
import java.sql.*;
import java.util.Arrays;

/**
 * Created by wanggq on 2017/9/26.
 */
@Component
public class Excel07Export {
    public static final String rootPath = "f:/excel";
    private Logger logger = Logger.getLogger(this.getClass().getName());

//    public void export(Connection con, String sql, ParamProcessor pp, Workbook wb) throws SQLException {
//        ResultSet rs = null;
//        try (PreparedStatement ps = con.prepareStatement(sql)) {
//            pp.build(ps);
//            rs = ps.executeQuery();
//            while (rs.next()) {
//                logger.info(rs.getString(1));
//            }
//        } finally {
//            closeResultSet(rs);
//        }
//    }

    public void export(Workbook workbook) throws TooManyColumnsException {
        exportTo(workbook, rootPath);
    }

    public void exportTo(Workbook workbook, String rootPath) throws TooManyColumnsException {
        File file = new File(rootPath, workbook.getName());
        if (!file.exists()) {
            if (!file.mkdir()) {
                logger.error("创建文件失败.");
                return;
            }
        } else {
            boolean removeSelf = false;
            FileUtil.rmRf(file, removeSelf);
        }

//        String path = file.getPath() + "/xl";


        Sheet[] sheets = workbook.getSheets();
        int n;
        for (int i = 0; i < sheets.length; i++) {
            Sheet sheet = sheets[i];
            if ((n = sheet.getHeadColumns().length) > Const.Limit.MAX_COLUMNS_ON_SHEET) {
                throw new TooManyColumnsException(n);
            }
            if (sheet.getAutoSize() == 0) {
                if (workbook.isAutoSize()) {
                    sheet.autoSize();
                } else {
                    sheet.fixSize();
                }
            }
            sheet.setId(i + 1);
        }

        File xl = new File(file, "xl");
        if (!xl.exists()) {
            xl.mkdirs();
        }

        // 最先做水印, 写各sheet时需要使用
        workbook.madeMark(xl);

        // 写各worksheet数字
        Arrays.asList(sheets).stream().forEach(e -> {
            e.writeTo(file);
            e.close();
        });

        // Write SharedString, Styles and workbook.xml
        workbook.writeXML(new File(file, "xl"));

        // zip
        File zipFile = new ZipCompressor(file.getPath()).compressSubs(file);
        // 如果文件存在则在文件名后加下标
        File o = new File(file.getPath() + Const.Suffix.EXCEL_07);
        if (o.exists()) {
            final String fname = file.getName();
            String[] os = o.getParentFile().list((dir, name) ->
                    new File(dir, name).isFile()
                            && name.startsWith(fname)
                            && name.endsWith(Const.Suffix.EXCEL_07)
            );
            int len = os.length;
            String new_name;
            do {
                new_name = fname + " (" + len++ + ")" + Const.Suffix.EXCEL_07;
                n = StringUtil.indexOf(os, new_name);
            } while (n > -1);
            o = new File(file.getParent(), new_name);
        }
        // Rename to xlsx
        if (!zipFile.renameTo(o)) {
            // TODO rename file
        }

        // Delete src file
        boolean delSelf = true;
//        FileUtil.rmRf(file, delSelf);
    }

    protected void closeResultSet(ResultSet rs) {
        if (rs != null) {
            try {
                rs.close();
            } catch (SQLException e) {
            }
        }
    }


}

package net.cua.export;

import net.cua.export.entity.TooManyColumnsException;
import net.cua.export.manager.Const;
import net.cua.export.entity.e7.Sheet;
import net.cua.export.entity.e7.Workbook;
import net.cua.export.util.FileUtil;
import net.cua.export.util.StringUtil;
import net.cua.export.util.ZipCompressor;
import org.apache.log4j.Logger;

import java.io.*;
import java.nio.file.*;
import java.nio.file.attribute.PosixFilePermissions;

/**
 * https://msdn.microsoft.com/library
 * https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet(v=office.14).aspx#
 * Created by guanquan.wang at 2017/9/26.
 */
public class Excelx {
    private static Logger logger = Logger.getLogger(Excelx.class.getName());

    public static File create(Workbook workbook) throws IOException {
        try {
            return createTempZip(workbook).toFile();
        } catch (TooManyColumnsException e) {
            logger.error(e);
        }
        return null;
    }

    public static InputStream createToInputStream(Workbook workbook) throws IOException {
        try {
            return Files.newInputStream(createTempZip(workbook));
        } catch (TooManyColumnsException e) {
            logger.error(e);
        }
        return null;
    }

    public static void createTo(Workbook workbook, Path rootPath) throws IOException {
        if (!Files.exists(rootPath)) {
            Files.createDirectories(rootPath);
        }

        Path zip = null;
        try {
            zip = createTempZip(workbook);
        } catch (TooManyColumnsException e) {
            logger.error(e);
        }
        // 如果文件存在则在文件名后加下标
        Path o = Paths.get(rootPath.toString(), workbook.getName() + Const.Suffix.EXCEL_07);
        if (Files.exists(o)) {
            final String fname = workbook.getName();
            Path parent = o.getParent();
            if (parent != null && Files.exists(parent)) {
                String[] os = parent.toFile().list((dir, name) ->
                        new File(dir, name).isFile()
                                && name.startsWith(fname)
                                && name.endsWith(Const.Suffix.EXCEL_07)
                );
                String new_name;
                if (os != null) {
                    int len = os.length, n;
                    do {
                        new_name = fname + " (" + len++ + ")" + Const.Suffix.EXCEL_07;
                        n = StringUtil.indexOf(os, new_name);
                    } while (n > -1);
                } else {
                    new_name = fname + Const.Suffix.EXCEL_07;
                }
                o = Paths.get(parent.toString(), new_name);
            } else {
                // Rename to xlsx
                Files.move(zip, o, StandardCopyOption.REPLACE_EXISTING);
                return;
            }
        }
        // Rename to xlsx
        Files.move(zip, o);
    }

    public static void createTo(Workbook workbook, String rootPath) throws IOException {
        createTo(workbook, Paths.get(rootPath));
    }

    protected static Path createTempZip(Workbook workbook) throws IOException, TooManyColumnsException {
        Path root = Files.createTempDirectory("eec+"
                , PosixFilePermissions.asFileAttribute(PosixFilePermissions.fromString("rwxr-x---")));

        Path xl = Files.createDirectory(Paths.get(root.toString(), "xl"));
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

        // 最先做水印, 写各sheet时需要使用
        workbook.madeMark(xl);

        // 写各worksheet内容
        for (Sheet e : sheets) {
            e.writeTo(xl);
            if (e.getWaterMark() != null && e.getWaterMark().delete()); // Delete template image
        }

        // Write SharedString, Styles and workbook.xml
        workbook.writeXML(xl);
        if (workbook.getWaterMark() != null && workbook.getWaterMark().delete()); // Delete template image

        // zip
        File zipFile = new ZipCompressor(root.toString()).compressSubs(root.toFile());
        // Delete source files
        boolean delSelf = true;
        FileUtil.rmRf(root.toFile(), delSelf);

        return zipFile.toPath();
    }
}

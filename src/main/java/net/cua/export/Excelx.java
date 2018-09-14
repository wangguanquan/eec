package net.cua.export;

import net.cua.export.entity.TooManyColumnsException;
import net.cua.export.entity.e7.*;
import net.cua.export.manager.Const;
import net.cua.export.util.FileUtil;
import net.cua.export.util.StringUtil;
import net.cua.export.util.ZipUtil;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import java.io.*;
import java.nio.file.*;

/**
 * https://msdn.microsoft.com/library
 * https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet(v=office.14).aspx#
 * Created by guanquan.wang at 2017/9/26.
 */
public class Excelx {
//    static Logger logger = LogManager.getLogger(Excelx.class);
//    public static File create(Workbook workbook) throws IOException {
//        try {
//            return createTempZip(workbook).toFile();
//        } catch (TooManyColumnsException e) {
//            logger.error(e);
//        }
//        return null;
//    }
//
//    public static InputStream createToInputStream(Workbook workbook) throws IOException {
//        try {
//            return Files.newInputStream(createTempZip(workbook));
//        } catch (TooManyColumnsException e) {
//            logger.error(e);
//        }
//        return null;
//    }
//
//    public static void createTo(Workbook workbook, Path rootPath) throws IOException {
//        if (!Files.exists(rootPath)) {
//            FileUtil.mkdir(rootPath);
//        }
//
//        Path zip = null;
//        try {
//            zip = createTempZip(workbook);
//        } catch (TooManyColumnsException e) {
//            logger.error(e);
//        }
//
//        reMarkPath(zip, rootPath, workbook.getName());
//    }
//
//    protected static void reMarkPath(Path zip, Path path) throws IOException {
//        String str = path.toString(), name;
//        if (str.endsWith(Const.Suffix.EXCEL_07)) {
//            name = str.substring(str.lastIndexOf(Const.lineSeparator) + 1);
//        } else {
//            name = "新建文件";
//        }
//        reMarkPath(zip, path, name);
//    }
//
//    protected static void reMarkPath(Path zip, Path rootPath, String fileName) throws IOException {
//        // 如果文件存在则在文件名后加下标
//        Path o = rootPath.resolve(fileName + Const.Suffix.EXCEL_07);
//        if (Files.exists(o)) {
//            final String fname = fileName;
//            Path parent = o.getParent();
//            if (parent != null && Files.exists(parent)) {
//                String[] os = parent.toFile().list((dir, name) ->
//                        new File(dir, name).isFile()
//                                && name.startsWith(fname)
//                                && name.endsWith(Const.Suffix.EXCEL_07)
//                );
//                String new_name;
//                if (os != null) {
//                    int len = os.length, n;
//                    do {
//                        new_name = fname + " (" + len++ + ")" + Const.Suffix.EXCEL_07;
//                        n = StringUtil.indexOf(os, new_name);
//                    } while (n > -1);
//                } else {
//                    new_name = fname + Const.Suffix.EXCEL_07;
//                }
//                o = parent.resolve(new_name);
//            } else {
//                // Rename to xlsx
//                Files.move(zip, o, StandardCopyOption.REPLACE_EXISTING);
//                return;
//            }
//        }
//        // Rename to xlsx
//        Files.move(zip, o);
//    }
//
//    public static void createTo(Workbook workbook, String rootPath) throws IOException {
//        createTo(workbook, Paths.get(rootPath));
//    }
//
//    protected static Path createTempZip(Workbook workbook) throws IOException, TooManyColumnsException {
//        Path root = FileUtil.mktmp("eec+");
//
//        Path xl = Files.createDirectory(root.resolve("xl"));
//        Sheet[] sheets = workbook.getSheets();
//        int n;
//        for (int i = 0; i < sheets.length; i++) {
//            Sheet sheet = sheets[i];
//            if ((n = sheet.getHeadColumns().length) > Const.Limit.MAX_COLUMNS_ON_SHEET) {
//                throw new TooManyColumnsException(n);
//            }
//            if (sheet.getAutoSize() == 0) {
//                if (workbook.isAutoSize()) {
//                    sheet.autoSize();
//                } else {
//                    sheet.fixSize();
//                }
//            }
//            sheet.setId(i + 1);
//        }
//
//        // 最先做水印, 写各sheet时需要使用
//        workbook.madeMark(xl);
//
//        // 写各worksheet内容
//        for (Sheet e : sheets) {
//            e.writeTo(xl);
//            if (e.getWaterMark() != null && e.getWaterMark().delete()); // Delete template image
//        }
//
//        // Write SharedString, Styles and workbook.xml
//        workbook.writeXML(xl);
//        if (workbook.getWaterMark() != null && workbook.getWaterMark().delete()); // Delete template image
//
//        // Zip compress
//        boolean compressRoot = false;
//        Path zipFile = ZipUtil.zip(root, compressRoot, root);
//
//        // Delete source files
//        boolean delSelf = true;
//        FileUtil.rm_rf(root.toFile(), delSelf);
//
//        return zipFile;
//    }
//
//    /**
//     * 按template模版定义的格式输出
//     * @param stream template流
//     * @param o 替换内容 javabean or map
//     * @param rootPath 存储路径
//     */
//    public static Path createByTemplateTo(InputStream stream, Object o, Path rootPath) throws IOException {
//        Path zip = template(stream, o);
//        reMarkPath(zip, rootPath);
//        return rootPath;
//    }
//
//    protected static Path template(InputStream stream, Object o) throws IOException {
//        // Store template stream as zip file
//        Path temp = FileUtil.mktmp("eec+");
//        ZipUtil.unzip(stream, temp);
//
//        // Bind data
//        EmbedTemplate bt = new EmbedTemplate(temp);
//        if (bt.check()) { // Check files
//            bt.bind(o);
//        }
//
//        // Zip compress
//        boolean compressRoot = false;
//        Path zipFile = ZipUtil.zip(temp, compressRoot, temp);
//
//        // Delete source files
//        boolean delSelf = true;
//        FileUtil.rm_rf(temp.toFile(), delSelf);
//
//        return zipFile;
//    }
}

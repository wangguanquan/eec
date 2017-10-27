package net.cua.export.manager;

import net.cua.export.entity.e7.*;
import net.cua.export.manager.docProps.App;
import net.cua.export.manager.docProps.Core;
import net.cua.export.util.FileUtil;

import java.io.*;
import java.lang.reflect.InvocationTargetException;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

/**
 * Created by wanggq on 2017/9/29.
 */
public class WorkbookManager {

//    public WorkbookManager of(File file, Workbook wb, String creator) {
//        // Content type
//        ContentType contentType = new ContentType();
//        contentType.add(new ContentType.Default(Const.ContentType.RELATIONSHIP, "rels"));
//        contentType.add(new ContentType.Default(Const.ContentType.XML, "xml"));
//        contentType.add(new ContentType.Override(Const.ContentType.SHAREDSTRING, "/xl/sharedStrings.xml"));
//        contentType.add(new ContentType.Override(Const.ContentType.WORKBOOK, "/xl/workbook.xml"));
//        contentType.addRel(new Relationship("xl/workbook.xml", Const.Relationship.OFFICE_DOCUMENT));
//
//        // docProps
//        App app = new App();
//        app.setCompany("蜗牛数字有限公司");
//
//        int n = wb.getSheetSize();
//        List<String> titleParts = new ArrayList<>(n);
//        Sheet[] sheets = wb.getSheets();
//        for (int i = 0; i < n; i++) {
//            titleParts.add(sheets[i].getName());
//        }
//        app.setTitlePards(titleParts);
//
//        try {
//            app.writeTo(file.getPath() + "/docProps/app.xml");
//            contentType.add(new ContentType.Override(Const.ContentType.APP, "/docProps/app.xml"));
//            contentType.addRel(new Relationship("docProps/app.xml", Const.Relationship.APP));
//        } catch (IllegalAccessException e) {
//            e.printStackTrace();
//        } catch (NoSuchMethodException e) {
//            e.printStackTrace();
//        } catch (InvocationTargetException e) {
//            e.printStackTrace();
//        }
//
//        Core core = new Core();
//        core.setCreator(creator);
//        core.setTitle(wb.getName());
////        core.setDescription("this is a test file");
//
////        core.setVersion("1.0");
////        core.setCategory("九阴; 手游; 点击");
//
//        core.setModified(new Date());
//
//        try {
//            core.writeTo(file.getPath() + "/docProps/core.xml");
//            contentType.add(new ContentType.Override(Const.ContentType.CORE, "/docProps/core.xml"));
//            contentType.addRel(new Relationship("docProps/core.xml", Const.Relationship.CORE));
//        } catch (IllegalAccessException e) {
//            e.printStackTrace();
//        } catch (NoSuchMethodException e) {
//            e.printStackTrace();
//        } catch (InvocationTargetException e) {
//            e.printStackTrace();
//        }
//
//
//        File xl = new File(file, "xl");
//        if (!xl.exists()) {
//            xl.mkdirs();
//        }
//
//        // theme
//        File themeP = new File(xl, "theme");
//        if (!themeP.exists()) {
//            themeP.mkdirs();
//        }
//
//        ClassLoader cl = getClass().getClassLoader();
//        FileUtil.copyFile(cl.getResourceAsStream("template/theme1.xml"), new File(themeP, "theme1.xml"));
//        wb.addRel(new Relationship("theme/theme1.xml", Const.Relationship.THEME));
//        contentType.add(new ContentType.Override(Const.ContentType.THEME, "/xl/theme/theme1.xml"));
//
//        // style
////        File styleFile = new File(xl, "styles.xml");
//        wb.addRel(new Relationship("styles.xml", Const.Relationship.STYLE));
//        contentType.add(new ContentType.Override(Const.ContentType.STYLE, "/xl/styles.xml"));
//
//        wb.addRel(new Relationship("sharedStrings.xml", Const.Relationship.SHARED_STRING));
//
//        if (wb.getWaterMark() != null && !wb.getWaterMark().isEmpty()) {
//            contentType.add(new ContentType.Default(Const.ContentType.PNG, "png"));
//        }
//
//        // TODO write sheet content type
//        for (int i = 0; i < wb.getSheetSize(); i++) {
//            contentType.add(new ContentType.Override(Const.ContentType.SHEET, "/xl/worksheets/sheet" + wb.getSheets()[i].getId() + ".xml"));
//        } // END
//
//
//        wb.writeXML(xl);
//
//        // write content type
//        contentType.wirte(file);
//
//        return null;
//    }


}

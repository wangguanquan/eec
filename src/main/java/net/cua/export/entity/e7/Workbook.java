package net.cua.export.entity.e7;

import net.cua.export.annotation.TopNS;
import net.cua.export.manager.Const;
import net.cua.export.manager.RelManager;
import net.cua.export.manager.docProps.App;
import net.cua.export.manager.docProps.Core;
import net.cua.export.processor.ParamProcessor;
import net.cua.export.util.DateUtil;
import net.cua.export.util.FileUtil;
import net.cua.export.util.StringUtil;
import org.dom4j.*;

import javax.imageio.ImageIO;
import java.awt.AlphaComposite;
import java.awt.BasicStroke;
import java.awt.Color;
import java.awt.Graphics2D;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.List;

/**
 * Created by wanggq on 2017/9/26.
 */
@TopNS(prefix = {"", "r"}, value = "workbook"
        , uri = {Const.SCHEMA_MAIN, Const.Relationship.RELATIONSHIP})
public class Workbook {
    private String name;
    private Sheet[] sheets;
    private String waterMark;
    private volatile int size;
    private Connection con;
    private RelManager relManager;
    private boolean autoSize;
    private String creator;

    private SharedStrings sst; // 共享字符区
    private Styles styles; // 共享样式

    public Workbook() {
        this(null);
    }

    public Workbook(String name) {
        this(name, null);
    }

    public Workbook(String name, String creator) {
        this.name = name;
        this.creator = creator;
//        this.waterMark = DateUtil.getToday() + " " + creator;
        sheets = new Sheet[3]; // 默认建3个sheet页
        relManager = new RelManager();

        sst = new SharedStrings();
        // Load styles
        styles = new Styles().load(getClass().getClassLoader().getResourceAsStream("template/styles.xml"));
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public final Sheet[] getSheets() {
        return Arrays.copyOf(sheets, size);
    }

    public void setSheets(final Sheet[] sheets) {
        this.sheets = sheets.clone();
    }

    public String getWaterMark() {
        return waterMark;
    }

    public void setWaterMark(String waterMark) {
        this.waterMark = waterMark;
    }

    public void addRel(Relationship rel) {
        relManager.add(rel);
    }

    public void setCon(Connection con) {
        this.con = con;
    }

//    public int getSheetSize() {
//        return size;
//    }

    public Workbook setAutoSize(boolean autoSize) {
        this.autoSize = autoSize;
        return this;
    }

    public boolean isAutoSize() {
        return autoSize;
    }

    public SharedStrings getSst() {
        return sst;
    }

    public Styles getStyles() {
        return styles;
    }

    public void setCreator(String creator) {
        this.creator = creator;
//        this.waterMark = DateUtil.getToday() + " " + creator;
    }

//    public Workbook addSheet(Sheet sheet) {
//        if (size >= sheets.length) {
//            sheets = Arrays.copyOf(sheets, size + 1);
//        }
//        sheets[size++] = sheet;
//        return this;
//    }
//
//    public Workbook addSheet(String name, Sheet.HeadColumn ... headColumns) {
//        if (size > sheets.length) {
//            sheets = Arrays.copyOf(sheets, size + 1);
//        }
//        Sheet sheet = new Sheet(name, headColumns);
//        sheets[size++] = sheet;
//
//        return this;
//    }

    public Sheet addSheet(String name, String sql, Sheet.HeadColumn ... headColumns) {
        int _size = size;
        if (_size >= sheets.length) {
            sheets = Arrays.copyOf(sheets, _size + 1);
        }
        StatementSheet sheet = new StatementSheet(this, name, headColumns);
        PreparedStatement ps = null;
        try {
            ps = con.prepareStatement(sql, ResultSet.TYPE_FORWARD_ONLY, ResultSet.CONCUR_READ_ONLY);
        } catch (SQLException e) {
            e.printStackTrace();
        }
        sheet.setPs(ps);
        sheets[_size] = sheet;
        synchronized (this) {
            size++;
        }
        return sheet;
    }

    public Sheet addSheet(String name, String sql, ParamProcessor pp, Sheet.HeadColumn ... headColumns) {
        int _size = size;
        if (_size >= sheets.length) {
            sheets = Arrays.copyOf(sheets, _size + 1);
        }
        StatementSheet sheet = new StatementSheet(this, name, headColumns);
        PreparedStatement ps = null;
        try {
            ps = con.prepareStatement(sql, ResultSet.TYPE_FORWARD_ONLY, ResultSet.CONCUR_READ_ONLY);
            pp.build(ps);
        } catch (SQLException e) {
            e.printStackTrace();
        }
        sheet.setPs(ps);
        sheets[_size] = sheet;

        synchronized (this) {
            size++;
        }
        return sheet;
    }

    public void insertSheet(int index, Sheet sheet) {
        int _size = size;
        if (_size >= sheets.length) {
            sheets = Arrays.copyOf(sheets, _size + 1);
        }

        synchronized (this) {
            if (sheets[index] == null) {
                sheets[index] = sheet;
            } else {
                for ( ; _size >= index; ) {
                    sheets[_size] = sheets[--_size];
                }
                sheets[index] = sheet;
            }
            size++;
        }
    }

    public void writeXML(File root) {
        // 水印
//        madeMark(root);

        // Content type
        ContentType contentType = new ContentType();
        contentType.add(new ContentType.Default(Const.ContentType.RELATIONSHIP, "rels"));
        contentType.add(new ContentType.Default(Const.ContentType.XML, "xml"));
        contentType.add(new ContentType.Override(Const.ContentType.SHAREDSTRING, "/xl/sharedStrings.xml"));
        contentType.add(new ContentType.Override(Const.ContentType.WORKBOOK, "/xl/workbook.xml"));
        contentType.addRel(new Relationship("xl/workbook.xml", Const.Relationship.OFFICE_DOCUMENT));

        // docProps
        App app = new App();
        app.setCompany("蜗牛数字有限公司");

        List<String> titleParts = new ArrayList<>(size);
        for (int i = 0; i < size; i++) {
            titleParts.add(sheets[i].getName());
            addRel(new Relationship("worksheets/sheet" + sheets[i].getId() + ".xml", Const.Relationship.SHEET));
        }
        app.setTitlePards(titleParts);

        try {
            app.writeTo(root.getParent() + "/docProps/app.xml");
            contentType.add(new ContentType.Override(Const.ContentType.APP, "/docProps/app.xml"));
            contentType.addRel(new Relationship("docProps/app.xml", Const.Relationship.APP));
        } catch (IllegalAccessException e) {
            e.printStackTrace();
        } catch (NoSuchMethodException e) {
            e.printStackTrace();
        } catch (InvocationTargetException e) {
            e.printStackTrace();
        }

        Core core = new Core();
        if (StringUtil.isNotEmpty(creator)) {
            core.setCreator(creator);
        } else {
            core.setCreator("guanquan.wang@yandex.com");
        }
        core.setTitle(name);
//        core.setDescription("this is a test file");

//        core.setVersion("1.0");
//        core.setCategory("九阴; 手游; 点击");

        core.setModified(new Date());

        try {
            core.writeTo(root.getParent() + "/docProps/core.xml");
            contentType.add(new ContentType.Override(Const.ContentType.CORE, "/docProps/core.xml"));
            contentType.addRel(new Relationship("docProps/core.xml", Const.Relationship.CORE));
        } catch (IllegalAccessException e) {
            e.printStackTrace();
        } catch (NoSuchMethodException e) {
            e.printStackTrace();
        } catch (InvocationTargetException e) {
            e.printStackTrace();
        }


//        File xl = new File(file, "xl");
//        if (!xl.exists()) {
//            xl.mkdirs();
//        }

        // theme
        File themeP = new File(root, "theme");
        if (!themeP.exists()) {
            themeP.mkdirs();
        }

        FileUtil.copyFile(getClass().getClassLoader().getResourceAsStream("template/theme1.xml"), new File(themeP, "theme1.xml"));
        addRel(new Relationship("theme/theme1.xml", Const.Relationship.THEME));
        contentType.add(new ContentType.Override(Const.ContentType.THEME, "/xl/theme/theme1.xml"));

        // style
//        File styleFile = new File(xl, "styles.xml");
        addRel(new Relationship("styles.xml", Const.Relationship.STYLE));
        contentType.add(new ContentType.Override(Const.ContentType.STYLE, "/xl/styles.xml"));

        addRel(new Relationship("sharedStrings.xml", Const.Relationship.SHARED_STRING));

        boolean hasWaterMark = StringUtil.isNotEmpty(waterMark);
        for (int i = 0; i < size; i++) {
            if (StringUtil.isNotEmpty(sheets[i].getWaterMark())) {
                hasWaterMark = true;
                break;
            }
        }
        if (hasWaterMark) {
            contentType.add(new ContentType.Default(Const.ContentType.PNG, "png"));
        }

        // TODO write sheet content type
        for (int i = 0; i < size; i++) {
            contentType.add(new ContentType.Override(Const.ContentType.SHEET, "/xl/worksheets/sheet" + sheets[i].getId() + ".xml"));
        } // END


        // write content type
        contentType.wirte(root.getParentFile());

        // Relationship
        try {
            relManager.write(root, StringUtil.lowFirstKey(this.getClass().getSimpleName()) + Const.Suffix.XML);
        } catch (IllegalAccessException e) {
            e.printStackTrace();
        }

        // workbook.xml
        writeSelf(root);

        // styles
        styles.writeTo(new File(root, "styles.xml"));

        // share string
        sst.write(root);
    }

    public void madeMark(File parent) {
        Relationship supRel = null;
        if (waterMark != null && !waterMark.isEmpty()) {
            try {
                File temp = createWaterMark(waterMark);
                File descFile = new File(parent, "media");
                if (!descFile.exists()) {
                    descFile.mkdirs();
                }
                File picture = new File(descFile, "image1.png");
                FileUtil.copyFile(temp, picture);
                supRel = new Relationship("../media/" + picture.getName(), Const.Relationship.IMAGE);
            } catch(IOException e) {
                e.printStackTrace();
            }
        }
        String wm;
        for (int i = 0; i < size; i++) {
            if (StringUtil.isNotEmpty(wm = sheets[i].getWaterMark())) {
                try {
                    File temp = createWaterMark(wm);
                    File descFile = new File(parent, "media");
                    if (!descFile.exists()) {
                        descFile.mkdirs();
                    }
                    File picture = new File(descFile, "image1.png");
                    FileUtil.copyFile(temp, picture);
                    sheets[i].addRel(new Relationship("../media/" + picture.getName(), Const.Relationship.IMAGE));
                } catch(IOException e) {
                    e.printStackTrace();
                }
            } else if (StringUtil.isNotEmpty(waterMark)) {
                sheets[i].setWaterMark(waterMark);
                sheets[i].addRel(supRel);
            }
        }
    }

    /**
     * 生成水印图片
     *
     * @param watermark
     * @return
     * @throws IOException
     */
    public static File createWaterMark(String watermark) throws IOException {
        File outputFile = File.createTempFile("warterMark", "png");
        int width = 510; // 水印图片的宽度
        int height = 300; // 水印图片的高度 因为设置其他的高度会有黑线，所以拉高高度

        // 获取bufferedImage对象
        BufferedImage bi = new BufferedImage(width, height,
                BufferedImage.TYPE_INT_RGB);
        // 处理背景色，设置为 白色
        int minx = bi.getMinX();
        int miny = bi.getMinY();
        for (int i = minx; i < width; i++) {
            for (int j = miny; j < height; j++) {
                bi.setRGB(i, j, 0xffffff);
            }
        }

        // 获取Graphics2d对象
        Graphics2D g2d = bi.createGraphics();
        // 设置字体颜色为灰色
        g2d.setColor(new Color(240, 240, 240));
        // 设置图片的属性
        g2d.setStroke(new BasicStroke(1));
        // 设置字体
        g2d.setFont(new java.awt.Font("华文细黑", java.awt.Font.ITALIC, 50));
        // 设置字体倾斜度
        g2d.rotate(Math.toRadians(-10));

        // 写入水印文字 原定高度过小，所以累计写水印，增加高度
        for (int i = 1; i < 10; i++) {
            g2d.drawString(watermark, 0, 60 * i);
        }
        // 设置透明度
        g2d.setComposite(AlphaComposite.getInstance(AlphaComposite.SRC_OVER));
        // 释放对象
        g2d.dispose();
        ImageIO.write(bi, "png", outputFile);

        return outputFile;
    }


    protected void writeSelf(File root) {
        DocumentFactory factory = DocumentFactory.getInstance();
        //use the factory to create a root element
        Element rootElement = null;
        //use the factory to create a new document with the previously created root element
        boolean hasTopNs;
        String[] prefixs = null, uris = null;
        String rootName = null;
        TopNS topNs = getClass().getAnnotation(TopNS.class);
        if (hasTopNs = getClass().isAnnotationPresent(TopNS.class)) {
            prefixs = topNs.prefix();
            uris = topNs.uri();
            rootName = topNs.value();
            for (int i = 0; i < prefixs.length; i++) {
                if (prefixs[i].length() == 0) { // 创建前缀为空的命名空间
                    rootElement = factory.createElement(rootName, uris[i]);
                    break;
                }
            }
        }
        if (rootElement == null) {
            if (hasTopNs) {
                rootElement = factory.createElement(rootName);
            } else {
                // TODO echo error message
                return;
            }
        }

        if (hasTopNs) {
            for (int i = 0; i < prefixs.length; i++) {
                if (prefixs.length > 0) {
                    rootElement.add(Namespace.get(prefixs[i], uris[i]));
                }
            }
        }

        // book view
        rootElement.addElement("bookViews").addElement("workbookView").addAttribute("activeTab", "0");

        // sheets
        Element sheetEle = rootElement.addElement("sheets");
        for (int i = 0; i < size; i++) {
            Sheet sheetInfo = sheets[i];
            Element st = sheetEle.addElement(StringUtil.lowFirstKey(sheetInfo.getClass().getSuperclass().getSimpleName()))
                    .addAttribute("sheetId", String.valueOf(i + 1))
                    .addAttribute("name", sheetInfo.getName());
            Relationship rs = relManager.getByTarget("worksheets/sheet" + (i + 1) + Const.Suffix.XML);
            if (rs != null) {
                st.addAttribute(QName.get("id", Namespace.get("r", uris[StringUtil.indexOf(prefixs, "r")])), rs.getId());
            }
        }

        // Calculation Properties
        rootElement.addElement("calcPr").addAttribute("calcId", "124519");

        Document doc = factory.createDocument(rootElement);
        FileUtil.writeToDisk(doc, root.getPath() + "/" + rootName + Const.Suffix.XML); // write to desk
    }
}

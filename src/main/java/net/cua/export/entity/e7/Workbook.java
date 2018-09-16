package net.cua.export.entity.e7;

import net.cua.export.annotation.TopNS;
import net.cua.export.entity.ExportException;
import net.cua.export.entity.TooManyColumnsException;
import net.cua.export.entity.WaterMark;
import net.cua.export.entity.e7.style.Styles;
import net.cua.export.manager.Const;
import net.cua.export.manager.RelManager;
import net.cua.export.manager.docProps.App;
import net.cua.export.manager.docProps.Core;
import net.cua.export.processor.DownProcessor;
import net.cua.export.processor.ParamProcessor;
import net.cua.export.util.FileUtil;
import net.cua.export.util.StringUtil;
import net.cua.export.util.ZipUtil;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.dom4j.*;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.lang.reflect.InvocationTargetException;
import java.nio.file.*;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.*;

/**
 * https://poi.apache.org/encryption.html encrypted
 * https://msdn.microsoft.com/library
 * https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet(v=office.14).aspx#
 * Created by guanquan.wang on 2017/9/26.
 */
@TopNS(prefix = {"", "r"}, value = "workbook"
        , uri = {Const.SCHEMA_MAIN, Const.Relationship.RELATIONSHIP})
public class Workbook {
    Logger logger = LogManager.getLogger(getClass());

    private String name;
    private Sheet[] sheets;
    private WaterMark waterMark;
    private int size;
    private Connection con;
    private RelManager relManager;
    private boolean autoSize;
    private String creator, company;

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
        sheets = new Sheet[3]; // 默认建3个sheet页
        relManager = new RelManager();

        sst = new SharedStrings();
        // create styles
        styles = Styles.create();
    }

    public String getName() {
        return name;
    }

    public Workbook setName(String name) {
        this.name = name;
        return this;
    }

    public final Sheet[] getSheets() {
        return Arrays.copyOf(sheets, size);
    }

    public Workbook setSheets(final Sheet[] sheets) {
        this.sheets = sheets.clone();
        return this;
    }

    public WaterMark getWaterMark() {
        return waterMark;
    }

    public Workbook setWaterMark(WaterMark waterMark) {
        this.waterMark = waterMark;
        return this;
    }

    public Workbook addRel(Relationship rel) {
        relManager.add(rel);
        return this;
    }

    public Workbook setConnection(Connection con) {
        this.con = con;
        return this;
    }

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

    public Workbook setCreator(String creator) {
        this.creator = creator;
        return this;
    }

    public Workbook setCompany(String company) {
        this.company = company;
        return this;
    }

    public Workbook addSheet(Sheet sheet) {
        ensureCapacityInternal();
        sheets[size++] = sheet;
        return this;
    }

    public Workbook addSheet(String name, List<?> data, Sheet.Column ... columns) {
        int _size = size;
        Object o;
        if (data == null || data.isEmpty() || (o = getFirst(data)) == null) {
            ensureCapacityInternal();
            sheets[_size] = new EmptySheet(this, name, columns);
            return this;
        }

        int len = data.size(), limit = Const.Limit.MAX_ROWS_ON_SHEET_07 - 1, page = len / limit;
        if (len % limit > 0) {
            page++;
        }
        if (_size + page > sheets.length) {
            sheets = Arrays.copyOf(sheets, _size + page);
        }
        // 提前分页
        for (int i = 0, n; i < page; i++) {
            Sheet sheet;
            List<?> subList = data.subList(i * limit, (n = (i + 1) * limit) < len ? n : len);
            if (o instanceof Map) {
                sheet = new ListMapSheet(this, i > 0 ? name + " (" + i + ")" : name, columns).setData((List<Map<String, ?>>) subList);
            } else {
                sheet = new ListObjectSheet(this, i > 0 ? name + " (" + i + ")" : name, columns).setData(subList);
            }
            sheets[_size + i] = sheet;
        }
        size += page;
        return this;
    }

    public Object getFirst(List<?> data) {
        Object first = data.get(0);
        if (first != null) return first;
        int i = 1;
        do {
            first = data.get(i++);
        } while (first == null);
        return first;
    }

    public Workbook addSheet(String name, ResultSet rs, Sheet.Column ... columns) {
        ensureCapacityInternal();
        ResultSetSheet sheet = new ResultSetSheet(this, name, columns);
        sheet.setRs(rs);
        sheets[size++] = sheet;
        return this;
    }

    public Workbook addSheet(String name, String sql, Sheet.Column ... columns) throws SQLException {
        ensureCapacityInternal();
        StatementSheet sheet = new StatementSheet(this, name, columns);
        // TODO 提前分页
        PreparedStatement ps = con.prepareStatement(sql, ResultSet.TYPE_FORWARD_ONLY, ResultSet.CONCUR_READ_ONLY);
        sheet.setPs(ps);
        sheets[size++] = sheet;
        return this;
    }

    public Workbook addSheet(String name, String sql, ParamProcessor pp, Sheet.Column ... columns) throws SQLException {
        ensureCapacityInternal();
        StatementSheet sheet = new StatementSheet(this, name, columns);
        PreparedStatement ps = con.prepareStatement(sql, ResultSet.TYPE_FORWARD_ONLY, ResultSet.CONCUR_READ_ONLY);
        pp.build(ps);
        sheet.setPs(ps);
        sheets[size++] = sheet;
        return this;
    }

    public Workbook insertSheet(int index, Sheet sheet) {
        int _size = size;
        if (sheets[index] != null) {
            for ( ; _size > index; _size--) {
                sheets[_size] = sheets[_size - 1];
                sheets[_size].setId(sheets[_size].getId() + 1);
            }
        }
        sheets[index] = sheet;
        sheet.setId(index + 1);
        size++;
        return this;
    }

    public Workbook remove(int index) {
        if (index < 0 || index >= size) {
            return this;
        }
        if (index == size - 1) {
            sheets[index] = null;
        } else {
            for ( ; index < size - 1; index++) {
                sheets[index] = sheets[index + 1];
                sheets[index].setId(sheets[index].getId() - 1);
            }
        }
        size--;
        return this;
    }

    public Sheet getSheetAt(int index) {
        if (index < 0 || index >= size)
            throw new IndexOutOfBoundsException("Index: "+index+", Size: "+size);
        return sheets[index];
    }

    public Sheet getSheet(String sheetName) {
        for (Sheet sheet : sheets) {
            if (sheet.getName().equals(sheetName)) {
                return sheet;
            }
        }
        return null;
    }

    private void ensureCapacityInternal() {
        if (size >= sheets.length) {
            sheets = Arrays.copyOf(sheets, size + 1);
        }
    }

    public void writeXML(Path root) throws IOException, ExportException {

        // Content type
        ContentType contentType = new ContentType();
        contentType.add(new ContentType.Default(Const.ContentType.RELATIONSHIP, "rels"));
        contentType.add(new ContentType.Default(Const.ContentType.XML, "xml"));
        contentType.add(new ContentType.Override(Const.ContentType.SHAREDSTRING, "/xl/sharedStrings.xml"));
        contentType.add(new ContentType.Override(Const.ContentType.WORKBOOK, "/xl/workbook.xml"));
        contentType.addRel(new Relationship("xl/workbook.xml", Const.Relationship.OFFICE_DOCUMENT));

        // docProps
        App app = new App();
        if (StringUtil.isNotEmpty(company)) {
            app.setCompany(company);
        } else {
            app.setCompany("guanquan.wang@yandex.com");
        }

        List<String> titleParts = new ArrayList<>(size);
        for (int i = 0; i < size; i++) {
            titleParts.add(sheets[i].getName());
            addRel(new Relationship("worksheets/sheet" + sheets[i].getId() + Const.Suffix.XML, Const.Relationship.SHEET));
        }
        app.setTitlePards(titleParts);

        try {
            app.writeTo(root.getParent() + "/docProps/app.xml");
            contentType.add(new ContentType.Override(Const.ContentType.APP, "/docProps/app.xml"));
            contentType.addRel(new Relationship("docProps/app.xml", Const.Relationship.APP));
        } catch (IllegalAccessException | NoSuchMethodException | InvocationTargetException e) {
            throw new ExportException(e);
        }

        Core core = new Core();
        if (StringUtil.isNotEmpty(creator)) {
            core.setCreator(creator);
        } else {
            core.setCreator("guanquan.wang@yandex.com");
        }
        core.setTitle(name);

        core.setModified(new Date());

        try {
            core.writeTo(root.getParent() + "/docProps/core.xml");
            contentType.add(new ContentType.Override(Const.ContentType.CORE, "/docProps/core.xml"));
            contentType.addRel(new Relationship("docProps/core.xml", Const.Relationship.CORE));
        } catch (IllegalAccessException | NoSuchMethodException | InvocationTargetException e) {
            throw new ExportException(e);
        }

        Path themeP = root.resolve("theme");
        if (!Files.exists(themeP)) {
            Files.createDirectory(themeP);
        }
        try {
            Files.copy(getClass().getClassLoader().getResourceAsStream("template/theme1.xml"), themeP.resolve("theme1.xml"));
        } catch (IOException e) {
            e.printStackTrace();
        }
//        FileUtil.copyFile(getClass().getClassLoader().getResourceAsStream("template/theme1.xml"), new File(themeP, "theme1.xml"));
        addRel(new Relationship("theme/theme1.xml", Const.Relationship.THEME));
        contentType.add(new ContentType.Override(Const.ContentType.THEME, "/xl/theme/theme1.xml"));

        // style
//        File styleFile = new File(xl, "styles.xml");
        addRel(new Relationship("styles.xml", Const.Relationship.STYLE));
        contentType.add(new ContentType.Override(Const.ContentType.STYLE, "/xl/styles.xml"));

        addRel(new Relationship("sharedStrings.xml", Const.Relationship.SHARED_STRING));

        if (waterMark != null) {
            contentType.add(new ContentType.Default(waterMark.getContentType(), waterMark.getSuffix().substring(1)));
        }
        for (int i = 0; i < size; i++) {
            WaterMark wm = sheets[i].getWaterMark();
            if (wm != null) {
                contentType.add(new ContentType.Default(wm.getContentType(), wm.getSuffix().substring(1)));
            }
        }

        for (int i = 0; i < size; i++) {
            contentType.add(new ContentType.Override(Const.ContentType.SHEET, "/xl/worksheets/sheet" + sheets[i].getId() + Const.Suffix.XML));
        } // END


        // write content type
        contentType.write(root.getParent());

        // Relationship
        relManager.write(root, StringUtil.lowFirstKey(this.getClass().getSimpleName()) + Const.Suffix.XML);

        // workbook.xml
        writeSelf(root);

        // styles
        styles.writeTo(root.resolve("styles.xml"));

        // share string
        sst.write(root);
    }

    public void madeMark(Path parent) throws IOException {
        Relationship supRel = null;
        int n = 1;
        if (waterMark != null) {
            Path media = parent.resolve("media");
            if (!Files.exists(media)) {
                Files.createDirectory(media);
            }
            Path image = media.resolve("image" + n++ + waterMark.getSuffix());

            Files.copy(waterMark.get(), image);
            supRel = new Relationship("../media/" + image.getFileName(), Const.Relationship.IMAGE);
        }
        WaterMark wm;
        for (int i = 0; i < size; i++) {
            if ((wm = sheets[i].getWaterMark()) != null) {
                Path media = parent.resolve("media");
                if (!Files.exists(media)) {
                    Files.createDirectory(media);
                }
                Path image = media.resolve("image" + n++ + wm.getSuffix());
                Files.copy(wm.get(), image);
                sheets[i].addRel(new Relationship("../media/" + image.getFileName(), Const.Relationship.IMAGE));
            } else if (waterMark != null) {
                sheets[i].setWaterMark(waterMark);
                sheets[i].addRel(supRel);
            }
        }
    }

    protected void writeSelf(Path root) throws IOException {
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
            if (sheetInfo.isHidden()) {
                st.addAttribute("state", "hidden");
            }
            Relationship rs = relManager.getByTarget("worksheets/sheet" + (i + 1) + Const.Suffix.XML);
            if (rs != null) {
                st.addAttribute(QName.get("id", Namespace.get("r", uris[StringUtil.indexOf(prefixs, "r")])), rs.getId());
            }
        }

        // Calculation Properties
        rootElement.addElement("calcPr").addAttribute("calcId", "124519");

        Document doc = factory.createDocument(rootElement);
        FileUtil.writeToDisk(doc, root.resolve(rootName + Const.Suffix.XML)); // write to desk
    }

    //////////////////////////////////////////////////////
    protected Path createTemp() throws IOException, ExportException {
        Sheet[] sheets = getSheets();
        int n;
        for (int i = 0; i < sheets.length; i++) {
            Sheet sheet = sheets[i];
            if ((n = sheet.getColumns().length) > Const.Limit.MAX_COLUMNS_ON_SHEET) {
                throw new TooManyColumnsException(n);
            }
            if (sheet.getAutoSize() == 0) {
                if (isAutoSize()) {
                    sheet.autoSize();
                } else {
                    sheet.fixSize();
                }
            }
            sheet.setId(i + 1);
        }

        Path root = null;
        try {
            root = FileUtil.mktmp("eec+");
            logger.info(root);

            Path xl = Files.createDirectory(root.resolve("xl"));
            // 最先做水印, 写各sheet时需要使用
            madeMark(xl);

            // 写各worksheet内容
            for (Sheet e : sheets) {
                e.writeTo(xl);
                if (e.getWaterMark() != null && e.getWaterMark().delete()) ; // Delete template image
            }

            // Write SharedString, Styles and workbook.xml
            writeXML(xl);
            if (getWaterMark() != null && getWaterMark().delete()) ; // Delete template image

            // Zip compress
            boolean compressRoot = false;
            Path zipFile = ZipUtil.zip(root, compressRoot, root);

            // TODO Delete source files
            boolean delSelf = true;
            FileUtil.rm_rf(root.toFile(), delSelf);
            return zipFile;
        } catch (IOException | ExportException e) {
            // remove temp path
            if (root != null) FileUtil.rm_rf(root);
            throw e;
        }
    }

    protected static void reMarkPath(Path zip, Path path) throws IOException {
        String str = path.toString(), name;
        if (str.endsWith(Const.Suffix.EXCEL_07)) {
            name = str.substring(str.lastIndexOf(Const.lineSeparator) + 1);
        } else {
            name = "新建文件";
        }
        reMarkPath(zip, path, name);
    }

    protected static void reMarkPath(Path zip, Path rootPath, String fileName) throws IOException {
        // 如果文件存在则在文件名后加下标
        Path o = rootPath.resolve(fileName + Const.Suffix.EXCEL_07);
        if (Files.exists(o)) {
            final String fname = fileName;
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
                o = parent.resolve(new_name);
            } else {
                // Rename to xlsx
                Files.move(zip, o, StandardCopyOption.REPLACE_EXISTING);
                return;
            }
        }
        // Rename to xlsx
        Files.move(zip, o);
    }

    //////////////////////////Print Out/////////////////////////////
    public void writeTo(Path path) throws IOException, ExportException {
        if (!Files.isDirectory(path)) {
            writeTo(path.toFile());
            return;
        }
        if (!Files.exists(path)) {
            FileUtil.mkdir(path);
        }
// TODO delete temp zip file
        Path zip = is == null ? createTemp() : template();

        reMarkPath(zip, path, getName());
    }

    public void writeTo(OutputStream os) throws IOException, ExportException {
        Path zip = is == null ? createTemp() : template();
        Files.copy(zip, os);
    }

    public void writeTo(File file) throws IOException, ExportException {
        Path zip = is == null ? createTemp() : template();
        FileUtil.cp(zip, file);
    }

    /**
     * return excel path
     * @return
     */
    public void create(DownProcessor processor) throws IOException, ExportException {
        Path zip = createTemp();
        processor.build(zip);
    }


    /////////////////////////////////Template///////////////////////////////////
    InputStream is;
    Object o;
    public Workbook withTemplate(InputStream is) {
        this.is = is;
        return this;
    }

    public Workbook withTemplate(InputStream is, Object o) {
        this.is = is;
        this.o = o;
        return this;
    }

    protected Path template() throws IOException {
        // Store template stream as zip file
        Path temp = FileUtil.mktmp("eec+");
        ZipUtil.unzip(is, temp);

        // Bind data
        EmbedTemplate bt = new EmbedTemplate(temp);
        if (bt.check()) { // Check files
            bt.bind(o);
        }

        // Zip compress
        boolean compressRoot = false;
        Path zipFile = ZipUtil.zip(temp, compressRoot, temp);

        // Delete source files
        boolean delSelf = true;
        FileUtil.rm_rf(temp.toFile(), delSelf);

        return zipFile;
    }
}

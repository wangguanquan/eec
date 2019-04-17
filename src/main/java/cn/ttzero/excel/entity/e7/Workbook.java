/*
 * Copyright (c) 2019, guanquan.wang@yandex.com All Rights Reserved.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

package cn.ttzero.excel.entity.e7;

import cn.ttzero.excel.entity.I18N;
import cn.ttzero.excel.entity.style.Fill;
import cn.ttzero.excel.entity.style.PatternType;
import cn.ttzero.excel.manager.Const;
import cn.ttzero.excel.manager.RelManager;
import cn.ttzero.excel.manager.docProps.App;
import cn.ttzero.excel.manager.docProps.Core;
import cn.ttzero.excel.processor.ParamProcessor;
import cn.ttzero.excel.annotation.TopNS;
import cn.ttzero.excel.entity.ExportException;
import cn.ttzero.excel.entity.TooManyColumnsException;
import cn.ttzero.excel.entity.WaterMark;
import cn.ttzero.excel.entity.style.Styles;
import cn.ttzero.excel.processor.DownProcessor;
import cn.ttzero.excel.processor.Watch;
import cn.ttzero.excel.util.FileUtil;
import cn.ttzero.excel.util.StringUtil;
import cn.ttzero.excel.util.ZipUtil;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.dom4j.*;

import javax.naming.OperationNotSupportedException;
import java.awt.Color;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.lang.reflect.InvocationTargetException;
import java.nio.file.*;
import java.sql.*;
import java.util.*;
import java.util.Date;
import java.util.List;

/**
 * 工作簿是Excel的基础单元，一个xlsx文件对应一个工作簿实例
 * 先设置属性和添加Sheet最后调writeTo方法执行写操作。
 * writeTo和create是一个终止语句，应该放置在未尾否则设置将不会被反应到最终的Excel文件中。
 * @link https://poi.apache.org/encryption.html encrypted
 * @link https://msdn.microsoft.com/library
 * @link https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet(v=office.14).aspx#
 * @link https://docs.microsoft.com/zh-cn/previous-versions/office/office-12/ms406049(v=office.12)
 *
 * Created by guanquan.wang on 2017/9/26.
 */
@TopNS(prefix = {"", "r"}, value = "workbook"
        , uri = {Const.SCHEMA_MAIN, Const.Relationship.RELATIONSHIP})
public class Workbook {
    private Logger logger = LogManager.getLogger(getClass());
    /** 工作薄名，最终反应到Excel文件名*/
    private String name;
    private Sheet[] sheets;
    private WaterMark waterMark; // workbook背景，应用于所有sheet页
    private int size;
    private Connection con;
    private RelManager relManager; // 关联管理
    private boolean autoSize; // 自动列宽
    private int autoOdd = 0; // 自动隔行变色
    private String creator, company; // 创建者，公司
    private Fill oddFill; // 偶数行的填充
    private Watch watch; // 观察者
    private I18N i18N;

    private SharedStrings sst; // 共享字符区
    private Styles styles; // 共享样式

    /**
     * 创建未命名工作簿
     */
    public Workbook() {
        this(null);
    }

    /**
     * 创建工作簿
     * 保存时以此名称为文件名
     * @param name 名称
     */
    public Workbook(String name) {
        this(name, null);
    }

    /**
     * 创建工作簿
     * @param name 名称
     * @param creator 作者
     */
    public Workbook(String name, String creator) {
        this.name = name;
        this.creator = creator;
        sheets = new Sheet[3]; // 默认建3个sheet页
        relManager = new RelManager();

        sst = new SharedStrings();
        i18N = new I18N();
        // create styles
        styles = Styles.create(i18N);
    }

    /**
     * 返回工作簿名称
     * @return name
     */
    public String getName() {
        return name;
    }

    /**
     * 设置工作簿名称
     * @param name 名称
     * @return 工作簿
     */
    public Workbook setName(String name) {
        this.name = name;
        return this;
    }

    /**
     * 返回所有Sheet页
     * @return Sheet数组
     */
    public final Sheet[] getSheets() {
        return Arrays.copyOf(sheets, size);
    }

    /**
     * 设置Sheet页
     * @param sheets Sheet数组
     * @return 工作簿
     */
    public Workbook setSheets(final Sheet[] sheets) {
        this.sheets = sheets.clone();
        return this;
    }

    /**
     * 返回水印
     * <p>标准excel文件中没有水印，所谓水印就是设置图片背景然后平铺以达到效果
     * ，此水印打印的时候并不会被打印。
     * </p>
     * @return 工作簿
     */
    public WaterMark getWaterMark() {
        return waterMark;
    }

    /**
     * 设置水印
     * <p>可使用<code>WaterMark.of()</code>创建水印</p>
     * @param waterMark 水印
     * @return 工作簿
     * @link {WaterMark#of}
     */
    public Workbook setWaterMark(WaterMark waterMark) {
        this.waterMark = waterMark;
        return this;
    }

    private Workbook addRel(Relationship rel) {
        relManager.add(rel);
        return this;
    }

    /**
     * 设置数据库连接
     * <p>工作簿内部不会主动关闭数据库连接，需要外部手动关闭，
     * 此eec内部产生的Statement和ResultSet会主动关闭</p>
     * @param con 连接
     * @return 工作簿
     */
    public Workbook setConnection(Connection con) {
        this.con = con;
        return this;
    }

    /**
     * 设置列宽自动调整
     * @param autoSize boolean value
     * @return 工作簿
     */
    public Workbook setAutoSize(boolean autoSize) {
        this.autoSize = autoSize;
        return this;
    }

    /**
     * 返回列宽是否自动调整
     * @return boolean value
     */
    public boolean isAutoSize() {
        return autoSize;
    }

    SharedStrings getSst() {
        return sst;
    }

    /**
     * 返回所有样式
     * @return Styles
     */
    public Styles getStyles() {
        return styles;
    }

    /**
     * 设置作者
     * @param creator 作者
     * @return 工作簿
     */
    public Workbook setCreator(String creator) {
        this.creator = creator;
        return this;
    }

    /**
     * 设置公司名
     * @param company 公司名
     * @return 工作簿
     */
    public Workbook setCompany(String company) {
        this.company = company;
        return this;
    }

    /**
     * 取消隔行变色
     * @return 工作簿
     */
    public Workbook cancelOddFill() {
        this.autoOdd = 1;
        return this;
    }
    /**
     * 设置隔行变色的背景色，默认为#e2edda
     * @param fill 偶数行背景色
     * @return 工作簿
     */
    public Workbook setOddFill(Fill fill) {
        this.oddFill = fill;
        return this;
    }

    /**
     * 尾部添加Sheet
     * @param sheet Sheet
     * @return 工作簿
     */
    public Workbook addSheet(Sheet sheet) {
        ensureCapacityInternal();
        sheets[size++] = sheet;
        return this;
    }

    /**
     * 尾部添加Sheet，默认名称
     * @param data 数据，Map数组/对象数组
     * @param columns 表头
     * @return 工作簿
     */
    public Workbook addSheet(List<?> data, Sheet.Column ... columns) {
        return addSheet(null, data, columns);
    }

    /**
     * 尾部添加Sheet
     * @param name 名称
     * @param data 数据，Map数组/对象数组
     * @param columns 表头
     * @return 工作簿
     */
    @SuppressWarnings("unchecked")
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
        if (StringUtil.isEmpty(name)) {
            name = "Sheet";
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

    Object getFirst(List<?> data) {
        Object first = data.get(0);
        if (first != null) return first;
        int i = 1;
        do {
            first = data.get(i++);
        } while (first == null);
        return first;
    }
    /**
     * 尾部添加Sheet，未命名
     * @param rs ResultSet
     * @param columns 表头
     * @return 工作簿
     */
    public Workbook addSheet(ResultSet rs, Sheet.Column ... columns) {
        return addSheet(null, rs, columns);
    }

    /**
     * 尾部添加Sheet
     * @param name 名称
     * @param rs ResultSet
     * @param columns 表头
     * @return 工作簿
     */
    public Workbook addSheet(String name, ResultSet rs, Sheet.Column ... columns) {
        ensureCapacityInternal();
        ResultSetSheet sheet = new ResultSetSheet(this, name, columns);
        sheet.setRs(rs);
        sheets[size++] = sheet;
        return this;
    }
    /**
     * 尾部添加Sheet，未命名
     * @param sql SQL文
     * @param columns 列头
     * @return 工作簿
     * @throws SQLException SQL异常
     */
    public Workbook addSheet(String sql, Sheet.Column ... columns) throws SQLException {
        return addSheet(null, sql, columns);
    }
    /**
     * 尾部添加Sheet
     * @param name 名称
     * @param sql SQL文
     * @param columns 列头
     * @return 工作簿
     * @throws SQLException SQL异常
     */
    public Workbook addSheet(String name, String sql, Sheet.Column ... columns) throws SQLException {
        ensureCapacityInternal();
        StatementSheet sheet = new StatementSheet(this, name, columns);
        PreparedStatement ps = con.prepareStatement(sql, ResultSet.TYPE_FORWARD_ONLY, ResultSet.CONCUR_READ_ONLY);
        ps.setFetchSize(Integer.MIN_VALUE);
        ps.setFetchDirection(ResultSet.FETCH_REVERSE);
        sheet.setPs(ps);
        sheets[size++] = sheet;
        return this;
    }
    /**
     * 尾部添加Sheet，未命令名
     * <p>采用jdbc方式设置SQL参数
     * eq: <code>workbook.addSheet("users", "select id, name from users where `class` = ?", ps -> ps.setString(1, "middle") ...</code>
     * </p>
     * @param sql SQL文
     * @param pp 参数
     * @param columns 列头
     * @return 工作簿
     * @throws SQLException SQL异常
     */
    public Workbook addSheet(String sql, ParamProcessor pp, Sheet.Column ... columns) throws SQLException {
        return addSheet(null, sql, pp, columns);
    }
    /**
     * 尾部添加Sheet
     * <p>采用jdbc方式设置SQL参数
     * eq: <code>workbook.addSheet("users", "select id, name from users where `class` = ?", ps -> ps.setString(1, "middle") ...</code>
     * </p>
     * @param name 名称
     * @param sql SQL文
     * @param pp 参数
     * @param columns 列头
     * @return 工作簿
     * @throws SQLException SQL异常
     */
    public Workbook addSheet(String name, String sql, ParamProcessor pp, Sheet.Column ... columns) throws SQLException {
        ensureCapacityInternal();
        StatementSheet sheet = new StatementSheet(this, name, columns);
        PreparedStatement ps = con.prepareStatement(sql, ResultSet.TYPE_FORWARD_ONLY, ResultSet.CONCUR_READ_ONLY);
        ps.setFetchSize(Integer.MIN_VALUE);
        ps.setFetchDirection(ResultSet.FETCH_REVERSE);
        pp.build(ps);
        sheet.setPs(ps);
        sheets[size++] = sheet;
        return this;
    }
    /**
     * 尾部添加Sheet，未命名
     * @param ps PreparedStatement
     * @param columns 列头
     * @return 工作簿
     * @throws SQLException SQL异常
     */
    public Workbook addSheet(PreparedStatement ps, Sheet.Column ... columns) throws SQLException {
        return addSheet(null, ps, columns);
    }
    /**
     * 尾部添加Sheet
     * @param name 名称
     * @param ps PreparedStatement
     * @param columns 列头
     * @return 工作簿
     * @throws SQLException SQL异常
     */
    public Workbook addSheet(String name, PreparedStatement ps, Sheet.Column ... columns) throws SQLException {
        ensureCapacityInternal();
        StatementSheet sheet = new StatementSheet(this, name, columns);
        ps.setFetchSize(Integer.MIN_VALUE);
        ps.setFetchDirection(ResultSet.FETCH_REVERSE);
        sheet.setPs(ps);
        sheets[size++] = sheet;
        return this;
    }
    /**
     * 尾部添加Sheet，未命令名
     * <p>采用jdbc方式设置SQL参数
     * eq: <code>workbook.addSheet("users", "select id, name from users where `class` = ?", ps -> ps.setString(1, "middle") ...</code>
     * </p>
     * @param ps PreparedStatement
     * @param pp 参数
     * @param columns 列头
     * @return 工作簿
     * @throws SQLException SQL异常
     */
    public Workbook addSheet(PreparedStatement ps, ParamProcessor pp, Sheet.Column ... columns) throws SQLException {
        return addSheet(null, ps, pp, columns);
    }
    /**
     * 尾部添加Sheet
     * <p>采用jdbc方式设置SQL参数
     * eq: <code>workbook.addSheet("users", "select id, name from users where `class` = ?", ps -> ps.setString(1, "middle") ...</code>
     * </p>
     * @param name 名称
     * @param ps PreparedStatement
     * @param pp 参数
     * @param columns 列头
     * @return 工作簿
     * @throws SQLException SQL异常
     */
    public Workbook addSheet(String name, PreparedStatement ps, ParamProcessor pp, Sheet.Column ... columns) throws SQLException {
        ensureCapacityInternal();
        StatementSheet sheet = new StatementSheet(this, name, columns);
        ps.setFetchSize(Integer.MIN_VALUE);
        ps.setFetchDirection(ResultSet.FETCH_REVERSE);
        pp.build(ps);
        sheet.setPs(ps);
        sheets[size++] = sheet;
        return this;
    }
    /**
     * 在指定位置上插入Sheet
     * @param index 下标从0开始
     * @param sheet 要插入的Sheet
     * @return 工作簿
     */
    public Workbook insertSheet(int index, Sheet sheet) {
        ensureCapacityInternal();
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

    /**
     * 移除指定下标的Sheet
     * @param index 下标从0开始
     * @return 工作簿
     */
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

    /**
     * 返回指定下标的Sheet
     * @param index 下标从0开始
     * @return Sheet
     */
    public Sheet getSheetAt(int index) {
        if (index < 0 || index >= size)
            throw new IndexOutOfBoundsException("Index: "+index+", Size: "+size);
        return sheets[index];
    }

    /**
     * 返回指定名称的Sheet
     * @param sheetName sheet name
     * @return Sheet, 未打到时返回null
     */
    public Sheet getSheet(String sheetName) {
        for (Sheet sheet : sheets) {
            if (sheet.getName().equals(sheetName)) {
                return sheet;
            }
        }
        return null;
    }

    /**
     * 添加观察者
     * @param watch 观察者
     * @return 工作簿
     */
    public Workbook watch(Watch watch) {
        this.watch = watch;
        return this;
    }

    /**
     * 以Excel 97-2003格式保存
     * @return 工作薄
     */
    public Workbook saveAsExcel2003() {
        throw new ExportException(new OperationNotSupportedException("Excel97-2003 Not support now."));
    }

    /**
     * output export info
     */
    void what(String code) {
        String msg = i18N.get(code);
        logger.debug(msg);
        if (watch != null) {
            watch.what(msg);
        }
    }

    /**
     * output export info
     */
    void what(String code, String ... args) {
        String msg = i18N.get(code, args);
        logger.debug(msg);
        if (watch != null) {
            watch.what(msg);
        }
    }

    private void ensureCapacityInternal() {
        if (size >= sheets.length) {
            sheets = Arrays.copyOf(sheets, size + 1);
        }
    }

    private void writeXML(Path root) throws IOException, ExportException {

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
            app.setCompany("cn.ttzero");
        }

        // Read app and version from pom
        try {
            Properties pom = new Properties();
            pom.load(getClass().getClassLoader().getResourceAsStream("META-INF/maven/cn.ttzero/eec/pom.properties"));
            app.setApplication(pom.getProperty("groupId") + '.' + pom.getProperty("artifactId"));
            app.setAppVersion(pom.getProperty("version"));
        } catch (IOException e) {
            // Nothing
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
        core.setCreated(new Date());
        if (StringUtil.isNotEmpty(creator)) {
            core.setCreator(creator);
        } else {
            core.setCreator(System.getProperty("user.name"));
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
            // Nothing
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

    private void madeMark(Path parent) throws IOException {
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

    private void writeSelf(Path root) throws IOException {
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
                what("9004", "workbook.xml");
                return;
            }
        }

        if (hasTopNs) {
            for (int i = 0; i < prefixs.length; i++) {
                rootElement.add(Namespace.get(prefixs[i], uris[i]));
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
        FileUtil.writeToDiskNoFormat(doc, root.resolve(rootName + Const.Suffix.XML)); // write to desk
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
            if (sheet.autoOdd == -1) {
                sheet.autoOdd = autoOdd;
            }
            // 默认隔行变色
            if (sheet.autoOdd == 0) {
                sheet.oddFill = styles.addFill(oddFill == null ? new Fill(PatternType.solid, new Color(226, 237, 218)) : oddFill);
            }
            sheet.setId(i + 1);
            // default worksheet name
            if (StringUtil.isEmpty(sheet.getName())) {
                sheet.setName("Sheet" + (i + 1));
            }
        }
        what("0001"); // 初始化完成

        Path root = null;
        try {
            root = FileUtil.mktmp("eec+"); // 创建临时文件
            what("0002", root.toString());

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
            if (waterMark != null && waterMark.delete()) ; // Delete template image
            what("0003");

            // Zip compress
            Path zipFile = ZipUtil.zipExcludeRoot(root, root);
            what("0004", zipFile.toString());

            // Delete source files
            FileUtil.rm_rf(root.toFile(), true);
            what("0005");
            return zipFile;
        } catch (IOException | ExportException e) {
            // remove temp path
            if (root != null) FileUtil.rm_rf(root);
            throw e;
        }
    }

    protected void reMarkPath(Path zip, Path path) throws IOException {
        String name;
        if (StringUtil.isEmpty(name = getName())) {
            name = i18N.getOrElse("no-name-file", "No name");
        }

        reMarkPath(zip, path, name);
    }

    protected void reMarkPath(Path zip, Path rootPath, String fileName) throws IOException {
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
        what("0006", o.toString());
    }

    //////////////////////////Print Out/////////////////////////////

    /**
     * 输出工作簿到指定文件夹下
     * <p>如果Path是文件夹，则将工作簿保存到该文件夹下，
     * 如果Path是文件，则将写到该文件下。
     * </p>
     * @param path 保存地址
     * @throws IOException IO异常
     * @throws ExportException 其它异常
     */
    public void writeTo(Path path) throws IOException, ExportException {
        if (!Files.isDirectory(path)) {
            writeTo(path.toFile());
            return;
        }
        if (!Files.exists(path)) {
            FileUtil.mkdir(path);
        }
        Path zip = is == null ? createTemp() : template();

        reMarkPath(zip, path);
    }

    /**
     * 输出到指定流
     * @param os OutputStream
     * @throws IOException IO异常
     * @throws ExportException 其它异常
     */
    public void writeTo(OutputStream os) throws IOException, ExportException {
        Path zip = is == null ? createTemp() : template();
        Files.copy(zip, os);
    }

    /**
     * 输出到文件
     * @param file 文件名
     * @throws IOException IO异常
     * @throws ExportException 其它异常
     */
    public void writeTo(File file) throws IOException, ExportException {
        Path zip = is == null ? createTemp() : template();
        FileUtil.cp(zip, file);
    }

    /**
     * 执行创建操作，然后再执行DownProcessor
     * eq;<code>workbook.create(path -> ... temp excel path)</code>
     * @param processor DownProcessor
     * @throws IOException IO异常
     * @throws ExportException 其它异常
     */
    public void create(DownProcessor processor) throws IOException, ExportException {
        Path zip = createTemp();
        processor.exec(zip);
    }


    /////////////////////////////////Template///////////////////////////////////
    private InputStream is;
    private Object o;

    /**
     * 设置模版
     * @param is 从流中读取模版
     * @return 工作簿
     */
    public Workbook withTemplate(InputStream is) {
        this.is = is;
        return this;
    }

    /**
     * 设置模版和绑定对象
     * @param is 从流中读取模版
     * @param o 绑定对象
     * @return 工作簿
     */
    public Workbook withTemplate(InputStream is, Object o) {
        this.is = is;
        this.o = o;
        return this;
    }

    protected Path template() throws IOException {
        what("0007");
        // Store template stream as zip file
        Path temp = FileUtil.mktmp("eec+");
        ZipUtil.unzip(is, temp);
        what("0008");

        // Bind data
        EmbedTemplate bt = new EmbedTemplate(temp, this);
        if (bt.check()) { // Check files
            bt.bind(o);
        }
        what("0003");

        // Zip compress
        Path zipFile = ZipUtil.zipExcludeRoot(temp, temp);
        what("0004", zipFile.toString());

        // Delete source files
        FileUtil.rm_rf(temp.toFile(), true);
        what("0005");

        return zipFile;
    }
}

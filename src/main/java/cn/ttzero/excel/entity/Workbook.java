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

package cn.ttzero.excel.entity;

import cn.ttzero.excel.entity.e7.*;
import cn.ttzero.excel.entity.style.Fill;
import cn.ttzero.excel.processor.ParamProcessor;
import cn.ttzero.excel.entity.style.Styles;
import cn.ttzero.excel.processor.Watch;
import cn.ttzero.excel.util.FileUtil;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import javax.naming.OperationNotSupportedException;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.lang.reflect.Constructor;
import java.nio.file.*;
import java.sql.*;
import java.util.*;
import java.util.List;

/**
 * 工作簿是Excel的基础单元，一个xlsx文件对应一个工作簿实例
 * 先设置属性和添加Sheet最后调writeTo方法执行写操作。
 * writeTo和create是一个终止语句，应该放置在未尾否则设置将不会被反应到最终的Excel文件中。
 *
 * @link https://poi.apache.org/encryption.html encrypted
 * @link https://msdn.microsoft.com/library
 * @link https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet(v=office.14).aspx#
 * @link https://docs.microsoft.com/zh-cn/previous-versions/office/office-12/ms406049(v=office.12)
 * <p>
 * Created by guanquan.wang on 2017/9/26.
 */
public class Workbook {
    private Logger logger = LogManager.getLogger(getClass());
    /**
     * 工作薄名，最终反应到Excel文件名
     */
    private String name;
    private Sheet[] sheets;
    private WaterMark waterMark; // workbook背景，应用于所有sheet页
    private int size;
    private Connection con;
    private boolean autoSize; // 自动列宽
    private int autoOdd = 0; // 自动隔行变色
    private String creator, company; // 创建者，公司
    private Fill oddFill; // 偶数行的填充
    private Watch watch; // A debug watcher
    private I18N i18N;

    private SharedStrings sst; // 共享字符区
    private Styles styles; // 共享样式

    private IWorkbookWriter workbookWriter;

    /**
     * 创建未命名工作簿
     */
    public Workbook() {
        this(null);
    }

    /**
     * 创建工作簿
     * 保存时以此名称为文件名
     *
     * @param name 名称
     */
    public Workbook(String name) {
        this(name, null);
    }

    /**
     * 创建工作簿
     *
     * @param name    名称
     * @param creator 作者
     */
    public Workbook(String name, String creator) {
        this.name = name;
        this.creator = creator;
        sheets = new Sheet[3]; // 默认建3个sheet页

        sst = new SharedStrings();
        i18N = new I18N();
        // create styles
        styles = Styles.create(i18N);

        // Default writer
        workbookWriter = new XMLWorkbookWriter(this);
    }

    /**
     * 返回工作簿名称
     *
     * @return name
     */
    public String getName() {
        return name;
    }

    /**
     * 设置工作簿名称
     *
     * @param name 名称
     * @return 工作簿
     */
    public Workbook setName(String name) {
        this.name = name;
        return this;
    }

    public int getAutoOdd() {
        return autoOdd;
    }

    public String getCreator() {
        return creator;
    }

    public String getCompany() {
        return company;
    }

    public Fill getOddFill() {
        return oddFill;
    }

    public I18N getI18N() {
        return i18N;
    }

    public int getSize() {
        return size;
    }

    public SharedStrings getSst() {
        return sst;
    }


    /**
     * 返回所有Sheet页
     *
     * @return Sheet数组
     */
    public final Sheet[] getSheets() {
        return Arrays.copyOf(sheets, size);
    }

    /**
     * 设置Sheet页
     *
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
     *
     * @return 工作簿
     */
    public WaterMark getWaterMark() {
        return waterMark;
    }

    /**
     * 设置水印
     * <p>可使用<code>WaterMark.of()</code>创建水印</p>
     *
     * @param waterMark 水印
     * @return 工作簿
     * @link {WaterMark#of}
     */
    public Workbook setWaterMark(WaterMark waterMark) {
        this.waterMark = waterMark;
        return this;
    }

    /**
     * 设置数据库连接
     * <p>工作簿内部不会主动关闭数据库连接，需要外部手动关闭，
     * 此eec内部产生的Statement和ResultSet会主动关闭</p>
     *
     * @param con 连接
     * @return 工作簿
     */
    public Workbook setConnection(Connection con) {
        this.con = con;
        return this;
    }

    /**
     * 设置列宽自动调整
     *
     * @param autoSize boolean value
     * @return 工作簿
     */
    public Workbook setAutoSize(boolean autoSize) {
        this.autoSize = autoSize;
        return this;
    }

    /**
     * 返回列宽是否自动调整
     *
     * @return boolean value
     */
    public boolean isAutoSize() {
        return autoSize;
    }


    /**
     * 返回所有样式
     *
     * @return Styles
     */
    public Styles getStyles() {
        return styles;
    }

    /**
     * 设置作者
     *
     * @param creator 作者
     * @return 工作簿
     */
    public Workbook setCreator(String creator) {
        this.creator = creator;
        return this;
    }

    /**
     * 设置公司名
     *
     * @param company 公司名
     * @return 工作簿
     */
    public Workbook setCompany(String company) {
        this.company = company;
        return this;
    }

    /**
     * 取消隔行变色
     *
     * @return 工作簿
     */
    public Workbook cancelOddFill() {
        this.autoOdd = 1;
        return this;
    }

    /**
     * 设置隔行变色的背景色，默认为#e2edda
     *
     * @param fill 偶数行背景色
     * @return 工作簿
     */
    public Workbook setOddFill(Fill fill) {
        this.oddFill = fill;
        return this;
    }

    /**
     * 尾部添加Sheet
     *
     * @param sheet Sheet
     * @return 工作簿
     */
    public Workbook addSheet(Sheet sheet) {
        ensureCapacityInternal();
        sheet.setWorkbook(this);
        sheets[size++] = sheet;
        return this;
    }

    /**
     * 尾部添加Sheet，默认名称
     *
     * @param data    数据，Map数组/对象数组
     * @param columns 表头
     * @return 工作簿
     */
    public Workbook addSheet(List<?> data, Sheet.Column... columns) {
        return addSheet(null, data, columns);
    }

    /**
     * 尾部添加Sheet
     *
     * @param name    名称
     * @param data    数据，Map数组/对象数组
     * @param columns 表头
     * @return 工作簿
     */
    @SuppressWarnings({"unchecked", "rawtypes"})
    public Workbook addSheet(String name, List<?> data, Sheet.Column... columns) {
        Object o;
        if (data == null || data.isEmpty() || (o = getFirst(data)) == null) {
            addSheet(new EmptySheet(name, columns));
            return this;
        }

        if (o instanceof Map) {
            addSheet(new ListMapSheet(name, columns).setData((List<Map<String, ?>>) data));
        } else {
            addSheet(new ListSheet(name, columns).setData(data));
        }
        return this;
    }

    private Object getFirst(List<?> data) {
        if (data == null) return null;
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
     *
     * @param rs      ResultSet
     * @param columns 表头
     * @return 工作簿
     */
    public Workbook addSheet(ResultSet rs, Sheet.Column... columns) {
        return addSheet(null, rs, columns);
    }

    /**
     * 尾部添加Sheet
     *
     * @param name    名称
     * @param rs      ResultSet
     * @param columns 表头
     * @return 工作簿
     */
    public Workbook addSheet(String name, ResultSet rs, Sheet.Column... columns) {
        ResultSetSheet sheet = new ResultSetSheet(name, columns);
        sheet.setRs(rs);
        addSheet(sheet);
        return this;
    }

    /**
     * 尾部添加Sheet，未命名
     *
     * @param sql     SQL文
     * @param columns 列头
     * @return 工作簿
     * @throws SQLException SQL异常
     */
    public Workbook addSheet(String sql, Sheet.Column... columns) throws SQLException {
        return addSheet(null, sql, columns);
    }

    /**
     * 尾部添加Sheet
     *
     * @param name    名称
     * @param sql     SQL文
     * @param columns 列头
     * @return 工作簿
     * @throws SQLException SQL异常
     */
    public Workbook addSheet(String name, String sql, Sheet.Column... columns) throws SQLException {
        StatementSheet sheet = new StatementSheet(name, columns);
        PreparedStatement ps = con.prepareStatement(sql, ResultSet.TYPE_FORWARD_ONLY, ResultSet.CONCUR_READ_ONLY);
        try {
            ps.setFetchSize(Integer.MIN_VALUE);
            ps.setFetchDirection(ResultSet.FETCH_REVERSE);
        } catch (SQLException e) {
            watch.what("Not support fetch size value of " + Integer.MIN_VALUE);
        }
        sheet.setPs(ps);
        addSheet(sheet);
        return this;
    }

    /**
     * 尾部添加Sheet，未命令名
     * <p>采用jdbc方式设置SQL参数
     * eq: <code>workbook.addSheet("users", "select id, name from users where `class` = ?", ps -> ps.setString(1, "middle") ...</code>
     * </p>
     *
     * @param sql     SQL文
     * @param pp      参数
     * @param columns 列头
     * @return 工作簿
     * @throws SQLException SQL异常
     */
    public Workbook addSheet(String sql, ParamProcessor pp, Sheet.Column... columns) throws SQLException {
        return addSheet(null, sql, pp, columns);
    }

    /**
     * 尾部添加Sheet
     * <p>采用jdbc方式设置SQL参数
     * eq: <code>workbook.addSheet("users", "select id, name from users where `class` = ?", ps -> ps.setString(1, "middle") ...</code>
     * </p>
     *
     * @param name    名称
     * @param sql     SQL文
     * @param pp      参数
     * @param columns 列头
     * @return 工作簿
     * @throws SQLException SQL异常
     */
    public Workbook addSheet(String name, String sql, ParamProcessor pp, Sheet.Column... columns) throws SQLException {
        StatementSheet sheet = new StatementSheet(name, columns);
        PreparedStatement ps = con.prepareStatement(sql, ResultSet.TYPE_FORWARD_ONLY, ResultSet.CONCUR_READ_ONLY);
        try {
            ps.setFetchSize(Integer.MIN_VALUE);
            ps.setFetchDirection(ResultSet.FETCH_REVERSE);
        } catch (SQLException e) {
            watch.what("Not support fetch size value of " + Integer.MIN_VALUE);
        }
        pp.build(ps);
        sheet.setPs(ps);
        addSheet(sheet);
        return this;
    }

    /**
     * 尾部添加Sheet，未命名
     *
     * @param ps      PreparedStatement
     * @param columns 列头
     * @return 工作簿
     * @throws SQLException SQL异常
     */
    public Workbook addSheet(PreparedStatement ps, Sheet.Column... columns) throws SQLException {
        return addSheet(null, ps, columns);
    }

    /**
     * 尾部添加Sheet
     *
     * @param name    名称
     * @param ps      PreparedStatement
     * @param columns 列头
     * @return 工作簿
     * @throws SQLException SQL异常
     */
    public Workbook addSheet(String name, PreparedStatement ps, Sheet.Column... columns) throws SQLException {
        StatementSheet sheet = new StatementSheet(name, columns);
        try {
            ps.setFetchSize(Integer.MIN_VALUE);
            ps.setFetchDirection(ResultSet.FETCH_REVERSE);
        } catch (SQLException e) {
            watch.what("Not support fetch size value of " + Integer.MIN_VALUE);
        }
        sheet.setPs(ps);
        addSheet(sheet);
        return this;
    }

    /**
     * 尾部添加Sheet，未命令名
     * <p>采用jdbc方式设置SQL参数
     * eq: <code>workbook.addSheet("users", "select id, name from users where `class` = ?", ps -> ps.setString(1, "middle") ...</code>
     * </p>
     *
     * @param ps      PreparedStatement
     * @param pp      参数
     * @param columns 列头
     * @return 工作簿
     * @throws SQLException SQL异常
     */
    public Workbook addSheet(PreparedStatement ps, ParamProcessor pp, Sheet.Column... columns) throws SQLException {
        return addSheet(null, ps, pp, columns);
    }

    /**
     * 尾部添加Sheet
     * <p>采用jdbc方式设置SQL参数
     * eq: <code>workbook.addSheet("users", "select id, name from users where `class` = ?", ps -> ps.setString(1, "middle") ...</code>
     * </p>
     *
     * @param name    名称
     * @param ps      PreparedStatement
     * @param pp      参数
     * @param columns 列头
     * @return 工作簿
     * @throws SQLException SQL异常
     */
    public Workbook addSheet(String name, PreparedStatement ps, ParamProcessor pp, Sheet.Column... columns) throws SQLException {
        ensureCapacityInternal();
        StatementSheet sheet = new StatementSheet(name, columns);
        try {
            ps.setFetchSize(Integer.MIN_VALUE);
            ps.setFetchDirection(ResultSet.FETCH_REVERSE);
        } catch (SQLException e) {
            watch.what("Not support fetch size value of " + Integer.MIN_VALUE);
        }
        pp.build(ps);
        sheet.setPs(ps);
        addSheet(sheet);
        return this;
    }

    /**
     * 在指定位置上插入Sheet
     *
     * @param index 下标从0开始
     * @param sheet 要插入的Sheet
     * @return 工作簿
     */
    public Workbook insertSheet(int index, Sheet sheet) {
        ensureCapacityInternal();
        int _size = size;
        if (sheets[index] != null) {
            for (; _size > index; _size--) {
                sheets[_size] = sheets[_size - 1];
                sheets[_size].setId(sheets[_size].getId() + 1);
            }
        }
        sheets[index] = sheet;
        sheet.setId(index + 1);
        sheet.setWorkbook(this);
        size++;
        return this;
    }

    /**
     * 移除指定下标的Sheet
     *
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
            for (; index < size - 1; index++) {
                sheets[index] = sheets[index + 1];
                sheets[index].setId(sheets[index].getId() - 1);
            }
        }
        size--;
        return this;
    }

    /**
     * 返回指定下标的Sheet
     *
     * @param index 下标从0开始
     * @return Sheet
     */
    public Sheet getSheetAt(int index) {
        if (index < 0 || index >= size)
            throw new IndexOutOfBoundsException("Index: " + index + ", Size: " + size);
        return sheets[index];
    }

    /**
     * 返回指定名称的Sheet
     *
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
     *
     * @param watch 观察者
     * @return 工作簿
     */
    public Workbook watch(Watch watch) {
        this.watch = watch;
        return this;
    }

    /**
     * 以Excel 97-2003格式保存
     *
     * @return 工作薄
     */
    public Workbook saveAsExcel2003() {
        try {
            Class<?> clazz = Class.forName("cn.ttzero.excel.entity.e3.BIFF8WorkbookWriter");
            Constructor<?> constructor = clazz.getDeclaredConstructor(this.getClass());
            workbookWriter = (IWorkbookWriter) constructor.newInstance(this);
        } catch (Exception e) {
            throw new ExcelWriteException(new OperationNotSupportedException("Excel97-2003 Not support now."));
        }
        return this;
    }

    /**
     * output export info
     */
    public void what(String code) {
        String msg = i18N.get(code);
        logger.debug(msg);
        if (watch != null) {
            watch.what(msg);
        }
    }

    /**
     * output export info
     */
    public void what(String code, String... args) {
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

    //////////////////////////Print Out/////////////////////////////

    /**
     * 输出工作簿到指定文件夹下
     * <p>如果Path是文件夹，则将工作簿保存到该文件夹下，
     * 如果Path是文件，则将写到该文件下。
     * </p>
     *
     * @param path 保存地址
     * @throws IOException         IO异常
     * @throws ExcelWriteException 其它异常
     */
    public void writeTo(Path path) throws IOException, ExcelWriteException {
        if (!Files.exists(path)) {
            String name = path.getFileName().toString();
            // write to file
            if (name.indexOf('.') > 0) {
                Path parent = path.getParent();
                FileUtil.mkdir(parent);
                writeTo(path.toFile());
                return;
                // write to directory
            } else FileUtil.mkdir(path);
        } else if (!Files.isDirectory(path)) {
            writeTo(path.toFile());
            return;
        }
        workbookWriter.writeTo(path);
    }

    /**
     * 输出到指定流
     *
     * @param os OutputStream
     * @throws IOException         IO异常
     * @throws ExcelWriteException 其它异常
     */
    public void writeTo(OutputStream os) throws IOException, ExcelWriteException {
        workbookWriter.writeTo(os);
    }

    /**
     * 输出到文件
     *
     * @param file 文件名
     * @throws IOException         IO异常
     * @throws ExcelWriteException 其它异常
     */
    public void writeTo(File file) throws IOException, ExcelWriteException {
        if (!file.getParentFile().exists()) {
            FileUtil.mkdir(file.toPath().getParent());
        }
        workbookWriter.writeTo(file);
    }

//    /**
//     * 执行创建操作，然后再执行DownProcessor
//     * eq;<code>workbook.create(path -> ... temp excel path)</code>
//     * @param processor DownProcessor
//     * @throws IOException IO异常
//     * @throws ExcelWriteException 其它异常
//     */
//    public void create(DownProcessor processor) throws IOException, ExcelWriteException {
//        Path zip = createTemp();
//        processor.exec(zip);
//    }


    /////////////////////////////////Template///////////////////////////////////
    private InputStream is;
    private Object o;

    public InputStream getTemplate() {
        return is;
    }

    public Object getBind() {
        return o;
    }

    /**
     * 设置模版
     *
     * @param is 从流中读取模版
     * @return 工作簿
     */
    public Workbook withTemplate(InputStream is) {
        this.is = is;
        return this;
    }

    /**
     * 设置模版和绑定对象
     *
     * @param is 从流中读取模版
     * @param o  绑定对象
     * @return 工作簿
     */
    public Workbook withTemplate(InputStream is, Object o) {
        this.is = is;
        this.o = o;
        return this;
    }

    protected Path template() throws IOException {
        return workbookWriter.template();
    }

    // --- Customize workbook writer
    public Workbook setWorkbookWriter(IWorkbookWriter workbookWriter) {
        this.workbookWriter = workbookWriter;
        this.workbookWriter.setWorkbook(this);
        return this;
    }
}

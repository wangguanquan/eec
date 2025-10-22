/*
 * Copyright (c) 2017-2018, guanquan.wang@hotmail.com All Rights Reserved.
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

package org.ttzero.excel.reader;

import org.dom4j.Document;
import org.dom4j.DocumentException;
import org.dom4j.Element;
import org.dom4j.Namespace;
import org.dom4j.QName;
import org.dom4j.io.SAXReader;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.ttzero.excel.entity.style.Theme;
import org.ttzero.excel.manager.NS;
import org.ttzero.excel.manager.TopNS;
import org.ttzero.excel.entity.IWorkbookWriter;
import org.ttzero.excel.entity.Relationship;
import org.ttzero.excel.entity.e7.ContentType;
import org.ttzero.excel.entity.style.Styles;
import org.ttzero.excel.manager.Const;
import org.ttzero.excel.manager.ExcelType;
import org.ttzero.excel.manager.RelManager;
import org.ttzero.excel.manager.docProps.App;
import org.ttzero.excel.manager.docProps.Core;
import org.ttzero.excel.manager.docProps.CustomProperties;
import org.ttzero.excel.util.DateUtil;
import org.ttzero.excel.util.FileUtil;
import org.ttzero.excel.util.SAXReaderUtil;
import org.ttzero.excel.util.StringUtil;

import java.io.Closeable;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.UncheckedIOException;
import java.lang.reflect.Constructor;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.Enumeration;
import java.util.HashSet;
import java.util.Iterator;
import java.util.List;
import java.util.Properties;
import java.util.Set;
import java.util.Spliterator;
import java.util.Spliterators;
import java.util.stream.Stream;
import java.util.stream.StreamSupport;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;

import static org.ttzero.excel.util.FileUtil.exists;
import static org.ttzero.excel.util.StringUtil.isEmpty;

/**
 * Excel读取工具
 *
 * <p>{@code ExcelReader}提供一组静态的{@link #read}方法，支持Iterator和Stream+Lambda读取xls和xlsx文件，
 * 你可以像操作集合类一样操作Excel。通过{@link Row#to}和{@link Row#too}方法可以将行数据转为指定对象，
 * 还可以使用{@link Row#toMap}方法转为LinkedHashMap，同时Row也提供更基础的类似于JDBC方式获取单元格的值。</p>
 *
 * <p>使用{@code ExcelReader}读取文件时不需要提前判断文件格式，Reader已内置类型判断并加载相应的解析器，
 * ExcelReader默认只能解析xlsx格式，如果需要解析xls则必须将{@code eec-e3-support}添加到classpath，它包含一个
 * {@code BIFF8Reader}用于解析BIFF8编码的xls格式文件。为保证功能统一几乎所有接口都由eec定义由support实现，
 * 大多数情况下ExcelReader和BIFF8Reader提供相同的功能，所以用{@code ExcelReader}读取excel文件时只需要一份代码</p>
 *
 * <p>读取过程中可能会产生一些临时文件，比如SharedString索引等临时文件，所以读取结束后需要关闭流并删除临时文件，
 * 建议使用{@code try...with...resource}块</p>
 *
 * <p>一个典型的读取示例如下：</p>
 * <pre>
 * try (ExcelReader reader = ExcelReader.read(path)) {
 *     // 读取所有工作表并打印
 *     reader.sheets().flatMap(Sheet::rows)
 *         .forEach(System.out::println);
 * } catch (IOException e) { }</pre>
 *
 * <p>参考文档:</p>
 * <p><a href="https://github.com/wangguanquan/eec/wiki/2-%E8%AF%BB%E5%8F%96Excel">读取Excel</a></p>
 * @author guanquan.wang on 2018-09-22
 */
public class ExcelReader implements Closeable {
    /**
     * LOGGER
     */
    protected static final Logger LOGGER = LoggerFactory.getLogger(ExcelReader.class);

    protected ExcelReader() { }

    /**
     * 保存所有工作表，读取工作表之前必须调用{@link Sheet#load}方法加载初始信息
     */
    protected Sheet[] sheets;
    /**
     * Excel临时文件路径，当传入参数为{@code InputStream}时会先将流写到此临时路径再进行读操作
     */
    private Path temp;
    /**
     * 文件格式，这里仅返回excel格式，其它文件一律返回{@code unknown}
     */
    private ExcelType type;
    /**
     * Excel文件基础信息包含作者、日期等信息，在windows操作系统上使用鼠标右键-&gt;属性-&gt;详细信息查看
     */
    private AppInfo appInfo;
    /**
     * 临时文件路径，读文件过程中产生的临时文件
     */
    protected Path tempDir;
    /**
     * 共享字符区
     */
    private SharedStrings sharedStringTable;

    /**
     * 图片管理器
     */
    protected Drawings drawings;
    /**
     * 全局的样式管理器
     */
    protected Styles styles;
    /**
     * Excel原始文件
     */
    protected ZipFile zipFile;

    /**
     * 以只读"值"的方式读取Excel文件，如果文件为{@code xls}格式则需要将{@code eec-e3-support}添加进classpath，未识别到文件类型则抛{@link ExcelReadException}
     *
     * @param path       excel文件路径
     * @return 一个Excel解析器 {@link ExcelReader}
     * @throws FileNotFoundException 文件不存在
     * @throws IOException           读取异常
     */
    public static ExcelReader read(Path path) throws IOException {
        if (!exists(path)) {
            throw new FileNotFoundException(path.toString());
        }
        // Check document type
        ExcelType type = getType(path);
        LOGGER.debug("File type: {}", type);
        ExcelReader reader;
        switch (type) {
            case XLSX:
                reader = new ExcelReader(path);
                break;
            case XLS:
                try {
                    Class<?> clazz = Class.forName("org.ttzero.excel.reader.BIFF8Reader");
                    Constructor<?> constructor = clazz.getDeclaredConstructor(Path.class);
                    reader = (ExcelReader) constructor.newInstance(path);
                } catch (ClassNotFoundException e) {
                    Properties pom = IWorkbookWriter.pom();
                    throw new ExcelReadException("Can not load 'org.ttzero.excel.reader.BIFF8Reader'."
                        + " Please add dependency [" + pom.getProperty("groupId") + ":eec-e3-support"
                        + ":" + pom.getProperty("version") + "] to parse excel 97~2003.", e);
                } catch (NoSuchMethodException | InstantiationException e) {
                    Properties pom = IWorkbookWriter.pom();
                    throw new ExcelReadException("It may be an exception caused by eec-e3-support version error."
                        + " Please add dependency [" + pom.getProperty("groupId") + ":eec-e3-support"
                        + ":" + pom.getProperty("version") + "]", e);
                } catch (IllegalAccessException | InvocationTargetException e) {
                    throw new ExcelReadException("Read excel failed.", e);
                }
                break;
            default:
                throw new ExcelReadException("Unknown file type.");
        }
        reader.type = type;
        return reader;
    }

    /**
     * 以只读"值"的方式读取Excel字节流
     *
     * @param stream     excel字节流
     * @return 一个Excel解析器 {@link ExcelReader}
     * @throws IOException 读取异常
     */
    public static ExcelReader read(InputStream stream) throws IOException {
        // 提前检查格式是否支持
        byte[] bytes = new byte[8];
        int n = stream.read(bytes);
        ExcelType type = typeOfStream(bytes, n);
        // 不是xls或xlsx格式
        if (type == ExcelType.UNKNOWN) throw new ExcelReadException("Unknown file type.");

        Path temp = Files.createTempFile("eec-", null);
        if (temp == null) throw new IOException("Create temp directory error. Please check your permission");
        OutputStream os = Files.newOutputStream(temp);
        os.write(bytes, 0, n);
        FileUtil.cp(stream, os);
        os.close();

        ExcelReader reader;
        try {
            reader = read(temp);
        } catch (IOException ex) {
            FileUtil.rm(temp);
            throw ex;
        } catch (Exception ex) {
            FileUtil.rm(temp);
            throw new IOException(ex);
        }
        reader.temp = temp;
        return reader;
    }

    /**
     * 获取当前Excel的文件类型，返回{@code xlsx}或{@code xls}，当文件不是excel时返回{@code unknown}
     *
     * @return {@link ExcelType}枚举
     */
    public ExcelType getType() {
        return type;
    }

    /**
     * 返回一个工作表的流，它将按顺序解析当前excel包含所有工作表（含隐藏工作表），
     * 此方法默认{@code load}工作表所以外部无需再次调用{@code load}方法
     *
     * @return 一个顺序的工作表流
     */
    public Stream<Sheet> sheets() {
        Iterator<Sheet> iter = new Iterator<Sheet>() {
            int n = 0;

            @Override
            public boolean hasNext() {
                return n < sheets.length;
            }

            @Override
            public Sheet next() {
                try {
                    // test and load sheet data
                    return sheets[n++].load();
                } catch (IOException e) {
                    throw new UncheckedIOException(e);
                }
            }
        };
        return StreamSupport.stream(Spliterators.spliterator(iter, sheets.length
            , Spliterator.ORDERED | Spliterator.NONNULL), false);
    }

    /**
     * 获取指定位置的工作表，此方法默认{@code load}工作表所以外部无需再次调用{@code load}方法
     *
     * @param index 工作表在excel的下标（从0开始）
     * @return 指定工作表，如果指定下标无工作表将抛{@code IndexOutOfBoundException}
     */
    public Sheet sheet(int index) {
        try {
            return sheets[index].load(); // lazy loading worksheet data
        } catch (IOException e) {
            throw new UncheckedIOException(e);
        }
    }

    /**
     * 获取指定名称的工作表，此方法默认{@code load}工作表所以外部无需再次调用{@code load}方法
     *
     * @param sheetName 工作表名
     * @return 指定工作表，如果不存在则返回{@code null}
     */
    public Sheet sheet(String sheetName) {
        try {
            for (Sheet t : sheets) {
                if (sheetName.equals(t.getName())) {
                    return t.load(); // lazy loading worksheet data
                }
            }
        } catch (IOException e) {
            throw new UncheckedIOException(e);
        }
        return null;
    }

    /**
     * 获取全部工作表，通过此方法获取的工作表在读取前需要先调用{@code load}方法
     *
     * @return 当前excel包含的所有工作表
     */
    public Sheet[] all() {
        return sheets;
    }

    /**
     * 获取当前excel包含的工作表数量
     *
     * @return 当前excel包含的工作表数量
     */
    public int getSheetCount() {
        return sheets != null ? sheets.length : 0;
    }

    /**
     * 关闭流并删除临时文件
     *
     * @throws IOException when fail close readers
     */
    @Override
    public void close() throws IOException {
        // Close all opened sheet
        if (sheets != null) {
            for (Sheet st : sheets) st.close();
        }

        // Close Shared String Table
        if (sharedStringTable != null) sharedStringTable.close();

        // Close source file
        if (zipFile != null) zipFile.close();

        // Remove temp file
        if (temp != null) FileUtil.rm(temp);

        // Remove temp dir
        if (tempDir != null) FileUtil.rm_rf(tempDir);
    }

    /**
     * Excel文件基础信息包含作者、日期等信息，在windows操作系统上使用鼠标右键-&gt;属性-&gt;详细信息查看
     *
     * @return {@link AppInfo}通过此对象可以获取excel详细属性
     */
    public AppInfo getAppInfo() {
        return appInfo != null ? appInfo : (appInfo = getGeneralInfo());
    }

    // --- PROTECTED FUNCTIONS

    /**
     * 配置必要的安全检查项，解析Excel文件之前会检查是否包含这些必须项，只要有一个不包含就抛{@link ExcelReadException}异常，
     * 可以在外部移除/添加检查项，当前支持的资源类型看请查看{@link Const.ContentType}
     */
    public static final Set<String> MUST_CHECK_PART = new HashSet<>(Arrays.asList(Const.ContentType.WORKBOOK
            , Const.ContentType.SHAREDSTRING, Const.ContentType.SHEET, Const.ContentType.STYLE));

    /**
     * 解析Content_Types并进行安全检查，必要安全检查不通过将抛{@link ExcelReadException}异常，
     * 必要检查项配置在{@link #MUST_CHECK_PART}中，外部可以视情况进行添加/移除
     *
     * @return 当前excel包含的全部资源类型
     */
    protected ContentType checkContentType() {
        // Read [Content_Types].xml
        ZipEntry entry = getEntry("[Content_Types].xml");
        if (entry == null) {
            if (temp != null) FileUtil.rm(temp);
            throw new ExcelReadException("The file format is incorrect or corrupted. [[Content_Types].xml]");
        }
        SAXReader reader = SAXReaderUtil.createDefault();
        Document document;
        try {
            document = reader.read(zipFile.getInputStream(entry));
        } catch (DocumentException | IOException e) {
            if (temp != null) FileUtil.rm(temp);
            throw new ExcelReadException("The file format is incorrect or corrupted. [[Content_Types].xml]");
        }
        ContentType contentType = new ContentType();
        List<Element> list = document.getRootElement().elements();
        for (Element e : list) {
            if ("Override".equals(e.getName())) {
                ContentType.Override override = new ContentType.Override(e.attributeValue("ContentType"), e.attributeValue("PartName"));
                entry = getEntry(override.getPartName());
                if (entry == null) {
                    if (MUST_CHECK_PART.contains(override.getContentType())) {
                        if (temp != null) FileUtil.rm(temp);
                        throw new ExcelReadException("The file format is incorrect or corrupted. [" + override.getPartName() + "]");
                    } else {
                        LOGGER.warn("{} is configured in [Content_Types].xml, but the corresponding file is missing.", override.getKey());
                    }
                }
                contentType.add(override);
            } else if ("Default".equals(e.getName())) {
                contentType.add(new ContentType.Default(e.attributeValue("ContentType"), e.attributeValue("Extension")));
            }
        }
        return contentType;
    }

    /**
     * 指定解析模式读取Excel文件
     *
     * @param stream     excel字节流
     * @throws IOException 读取异常
     */
    public ExcelReader(InputStream stream) throws IOException {
        // 提前检查格式是否支持
        byte[] bytes = new byte[8];
        int n = stream.read(bytes);
        ExcelType type = typeOfStream(bytes, n);
        // 不是xlsx格式
        if (type != ExcelType.XLSX) throw new ExcelReadException("Not a xlsx file.");

        Path temp = Files.createTempFile("eec-", null);
        if (temp == null) throw new IOException("Create temp directory error. Please check permission");
        OutputStream os = Files.newOutputStream(temp);
        os.write(bytes, 0, n);
        FileUtil.cp(stream, os);
        this.temp = temp;
        os.close();

        init(temp);
    }

    /**
     * 以只读"值"的方式读取指定路径的Excel文件
     *
     * @param path excel绝对路径
     * @throws IOException 读取异常
     */
    public ExcelReader(Path path) throws IOException {
        init(path);
    }

    /**
     * 初始化，初始化过程将进行内容检查，和创建全局属性（样式，字符共享区）以及工作表但不会实际读取工作表
     *
     * @param path       excel文件路径
     * @return 一个Excel解析器 {@link ExcelReader}
     * @throws IOException 读取异常
     */
    protected ExcelReader init(Path path) throws IOException {
        this.zipFile = new ZipFile(path.toFile());
        LOGGER.debug("Check file integrity.");

        // Check content-type
        ContentType contentType = checkContentType();
        if (contentType.hasDrawings()) {
            this.drawings = new XMLDrawings(this);
        }

//        // Check the file format and parse general information
//        appInfo = getGeneralInfo();

        // load workbook.xml
        SAXReader reader = SAXReaderUtil.createDefault();
        Document document;

        // Load SharedString
        ZipEntry entry = getEntry("xl/sharedStrings.xml");
        if (entry != null) {
            sharedStringTable = new SharedStrings(zipFile.getInputStream(entry), 0, 0).load();
        }

        // Load Styles
        entry = getEntry("xl/styles.xml");
        if (entry != null) {
            try {
                // Load Theme style first
                ZipEntry themeEntry = getEntry("xl/theme/theme1.xml");
                if (themeEntry != null) Theme.load(zipFile.getInputStream(themeEntry));

                // Then load custom styles
                styles = Styles.load(zipFile.getInputStream(entry));
            } catch (Exception ex) {
                LOGGER.warn("Parse style failed.", ex);
            }
        }
        // Construct a empty Styles
        if (styles == null) {
            styles = Styles.forReader();
        }

        entry = getEntry("xl/_rels/workbook.xml.rels");
        if (entry == null)
            throw new ExcelReadException("The file format is incorrect or corrupted. [xl/_rels/workbook.xml.rels]");

        try {
            document = reader.read(zipFile.getInputStream(entry));
        } catch (DocumentException | IOException e) {
            throw new ExcelReadException("The file format is incorrect or corrupted. [xl/_rels/workbook.xml.rels]");
        }

        List<Element> list = document.getRootElement().elements();
        Relationship[] rels = new Relationship[list.size()];
        int i = 0;
        for (Element e : list) {
            rels[i++] = new Relationship(e.attributeValue("Id"), e.attributeValue("Target"), e.attributeValue("Type"));
        }
        RelManager relManager = RelManager.of(rels);

        entry = getEntry("xl/workbook.xml");
        if (entry == null)
            throw new ExcelReadException("The file format is incorrect or corrupted. [xl/workbook.xml]");
        try {
            document = reader.read(zipFile.getInputStream(entry));
        } catch (DocumentException | IOException e) {
            throw new ExcelReadException("The file format is incorrect or corrupted. [xl/workbook.xml]");
        }
        Element root = document.getRootElement();
        Namespace ns = root.getNamespaceForPrefix("r");
        List<Sheet> sheets = new ArrayList<>();
        Iterator<Element> sheetIter = root.element("sheets").elementIterator();
        int index = 0;
        while (sheetIter.hasNext()) {
            Element e = sheetIter.next();
            XMLSheet sheet = (XMLSheet) sheetFactory();
            sheet.setName(e.attributeValue("name"));
            sheet.setId(Integer.parseInt(e.attributeValue("sheetId")));
            String state = e.attributeValue("state");
            sheet.setHidden("hidden".equals(state));
            Relationship r = relManager.getById(e.attributeValue(QName.get("id", ns)));
            if (r == null) {
                sheet.close();
                throw new ExcelReadException("The file format is incorrect or corrupted.");
            }
            String worksheetTarget = r.getTarget();
            if (!worksheetTarget.startsWith("worksheets")) {
                int a = worksheetTarget.indexOf("worksheets");
                if (a < 0) {
                    sheet.close();
                    throw new ExcelReadException("The file format is incorrect or corrupted.");
                }
                worksheetTarget = worksheetTarget.substring(a);
            }
            sheet.setPath("xl/" + worksheetTarget);
            entry = getEntry(sheet.path);
            if (entry == null) {
                sheet.close();
                throw new ExcelReadException("The file format is incorrect or corrupted.");
            }
            sheet.setZipFile(zipFile);
            sheet.setZipEntry(entry);
            // put shared string
            sheet.setSharedStrings(sharedStringTable);
            // Setting styles
            sheet.setStyles(styles);
            // Drawings
            sheet.setDrawings(drawings);
            sheet.setIndex(index++);
            sheets.add(sheet);
        }

        if (sheets.isEmpty())
            throw new ExcelReadException("The file format is incorrect or corrupted. [There has no worksheet]");

        Sheet[] sheets1 = new Sheet[sheets.size()];
        sheets.toArray(sheets1);

        this.sheets = sheets1;

        return this;
    }

    /**
     * 通过OPTION创建相应工作表
     *
     * @return Sheet extends XMLSheet
     */
    protected Sheet sheetFactory() {
        return new XMLSheet();
    }

    /**
     * 获取Shared String Table
     *
     * @return Shared String Table
     */
    public SharedStrings getSharedStrings() {
        return sharedStringTable;
    }

    /**
     * 判断文件格式，读取少量文件头字节来判断是否为BIFF和ZIP的文件签名
     *
     * @param path 临时文件路径
     * @return {@link ExcelType}枚举，非excel格式时返回{@link ExcelType#UNKNOWN}类型
     */
    public static ExcelType getType(Path path) {
        ExcelType type;
        try (InputStream is = Files.newInputStream(path)) {
            byte[] bytes = new byte[8];
            int len = is.read(bytes);
            type = typeOfStream(bytes, len);
        } catch (IOException e) {
            type = ExcelType.UNKNOWN;
        }
        return type;
    }

    // --- check
    private static ExcelType typeOfStream(byte[] bytes, int size) {
        ExcelType excelType = ExcelType.UNKNOWN;
        int length = Math.min(bytes.length, size);
        if (length < 4)
            return excelType;
        int type;
        type  = bytes[0]  & 0xFF;
        type += (bytes[1] & 0xFF) << 8;
        type += (bytes[2] & 0xFF) << 16;
        type += (bytes[3] & 0xFF) << 24;

        int zip = 0x04034B50;
        int b1  = 0xE011CFD0;
        int b2  = 0xE11AB1A1;

        if (type == zip) {
            excelType = ExcelType.XLSX;
        } else if (type == b1 && length >= 8) {
            type  = bytes[4]  & 0xFF;
            type += (bytes[5] & 0xFF) << 8;
            type += (bytes[6] & 0xFF) << 16;
            type += (bytes[7] & 0xFF) << 24;
            if (type == b2) excelType = ExcelType.XLS;
        }
        return excelType;
    }

    /**
     * 解析{@code docProps/app.xml}和{@code docProps/core.xml}文件获取文件基础信息，
     * 比如创建者、创建时间、分类等信息
     *
     * @return App信息
     */
    protected AppInfo getGeneralInfo() {
        // load app.xml
        SAXReader reader = SAXReaderUtil.createDefault();
        App app = null;
        Core core = null;
        ZipEntry entry = getEntry("docProps/app.xml");
        if (entry != null) {
            Document document = null;
            try {
                document = reader.read(zipFile.getInputStream(entry));
            } catch (DocumentException | IOException e) {
                LOGGER.warn("The file format is incorrect or corrupted. [docProps/app.xml]");
            }
            if (document != null) {
                Element root = document.getRootElement();
                app = new App();
                app.setCompany(root.elementText("Company"));
                app.setApplication(root.elementText("Application"));
                String v = root.elementText("AppVersion");
                if (StringUtil.isNotEmpty(v)) app.setAppVersion(v);
            }
        } else LOGGER.warn("The file format is incorrect or corrupted. [docProps/app.xml]");

        entry = getEntry("docProps/core.xml");
        if (entry != null) {
            Document document = null;
            try {
                document = reader.read(zipFile.getInputStream(entry));
            } catch (DocumentException | IOException e) {
                LOGGER.warn("The file format is incorrect or corrupted. [docProps/core.xml]");
            }
            if (document != null) {
                Element root = document.getRootElement();
                core = new Core();
                Class<Core> clazz = Core.class;
                TopNS topNS = clazz.getAnnotation(TopNS.class);
                String[] prefixs = topNS.prefix(), urls = topNS.uri();
                Field[] fields = clazz.getDeclaredFields();
                SimpleDateFormat format = DateUtil.utcDateTimeFormat.get();
                for (Field f : fields) {
                    NS ns = f.getAnnotation(NS.class);
                    if (ns == null) continue;

                    f.setAccessible(true);
                    int nsIndex = StringUtil.indexOf(prefixs, ns.value());
                    if (nsIndex < 0) continue;

                    Namespace namespace = new Namespace(ns.value(), urls[nsIndex]);
                    Class<?> type = f.getType();
                    String v = root.elementText(new QName(f.getName(), namespace));
                    if (isEmpty(v)) continue;
                    if (type == String.class) {
                        try {
                            f.set(core, v);
                        } catch (IllegalAccessException e) {
                            LOGGER.warn("Set field ({}) error.", f);
                        }
                    } else if (type == Date.class) {
                        try {
                            f.set(core, format.parse(v));
                        } catch (ParseException | IllegalAccessException e) {
                            LOGGER.warn("Set field ({}) error.", f);
                        }
                    }
                }
            }
        } else LOGGER.warn("The file format is incorrect or corrupted. [docProps/core.xml]");

        return new AppInfo(app, core);
    }

    /**
     * 将单元格坐标转为long类型，Excel单元格坐标由列+行组成如A1, B2等，
     * 转为long类型后第{@code 0-16}位为列号{@code 17-48}位为行号
     *
     * <blockquote><pre>
     * 单元格坐标    | 转换后long值
     * ------------+------------
     * A1          | 65537
     * AA10        | 655387
     * </pre></blockquote>
     *
     * @param r 单元格坐标
     * @return 转换后的值 高48位保存Row，低16位保存Col
     */
    public static long coordinateToLong(String r) {
        long v = 0L;
        int n = 0;
        for (int i = 0, len = r.length(); i < len; i++) {
            char value = r.charAt(i);
            if (value >= 'A' && value <= 'Z') {
                v = v * 26 + value - 'A' + 1;
            }
            else if (value >= '0' && value <= '9') {
                n = n * 10 + value - '0';
            }
            else if (value >= 'a' && value <= 'z') {
                v = v * 26 + value - 'a' + 1;
            }
            else
                throw new ExcelReadException("Column mark out of range: " + r);
        }
        return (v & 0x7FFF) | ((long) n) << 16;
    }

    /**
     * 获取Excel包含的所有图片，{@link Drawings.Picture}对象包含工作表的单元格行列信息，最重要的是包含{@code localPath}属性，
     * 它是图片的临时路径可以通过此路径复制图片
     *
     * @return 图片数组，如果不存在图片则返回{@code null}
     */
    public List<Drawings.Picture> listPictures() {
        return drawings != null ? drawings.listPictures() : null;
    }

    /**
     * 获取一个全局的样式对象 {@link Styles}
     *
     * @return 全局样式对象
     */
    public Styles getStyles() {
        return styles;
    }

    /**
     * 从压缩包中获取一个压缩文件
     *
     * @param name 压缩文件路径，必须是一个完整的路径
     * @return 如果实体存在则返回 {@link ZipEntry} 否则返回{@code null}
     */
    public ZipEntry getEntry(String name) {
        return getEntry(zipFile, toZipPath(name));
    }

    /**
     * 从压缩包中获取一个压缩文件字节流
     *
     * @param name 压缩文件路径，必须是一个完整的路径
     * @return 如果实体存在则返回该实体的{@code InputStream} 否则返回{@code null}
     * @throws IOException 读取异常
     */
    public InputStream getEntryStream(String name) throws IOException {
        ZipEntry entry = getEntry(zipFile, toZipPath(name));
        return entry != null ? zipFile.getInputStream(entry) : null;
    }

    /**
     * 从压缩包中获取一个压缩文件，为了兼容windows和linux系统的路径会进行{@code '/'}和{@code '\\'}
     * 两种分隔符匹配，如果路径无法匹配则遍历压缩包所有文件并忽略大小写匹配
     *
     * @param zipFile 压缩包
     * @param name    压缩文件路径，必须是一个完整的路径
     * @return 如果实体存在则返回 {@link ZipEntry} 否则返回{@code null}
     */
    public static ZipEntry getEntry(ZipFile zipFile, String name) {
        char c0 = name.charAt(0);
        if (c0 == '/' || c0 == '\\') name = name.substring(1);
        ZipEntry entry = zipFile.getEntry(name);
        // 如果原始路径查无则将路径替换为windows路径
        if (entry == null) entry = zipFile.getEntry(name.replace('/', '\\'));
        // 通过路径查无就遍历Zip包下所有资源忽略大小写匹配
        if (entry == null) {
            // Iterator entries
            Enumeration<? extends ZipEntry> entries = zipFile.entries();
            while (entries.hasMoreElements()) {
                ZipEntry e = entries.nextElement();
                String k = e.getName().replace('\\', '/');
                if (k.equalsIgnoreCase(name)) {
                    entry = e;
                    break;
                }
            }
        }
        return entry;
    }

    /**
     * 将string转换为zip允许的路径，将相对路径的前缀去掉
     *
     * @param path 实体路径
     * @return zip允许的路径
     */
    public static String toZipPath(String path) {
        int i = 0;
        if (path.startsWith("../") || path.startsWith("..\\")) i = 3;
        else if (path.startsWith("./") || path.startsWith(".\\")) i = 2;
        else if (path.charAt(0) == '/' || path.charAt(0) == '\\') i = 1;
        return i > 0 ? path.substring(i) : path;
    }

    /**
     * 获取自定义属性
     *
     * <p>返回数据类型说明，时间返回{@code java.util.Date}，布尔值返回{@code Boolean}，
     * 整数类型分情况返回{@code Integer}或{@code Long}，浮点数返回{@code BigDecimal}</p>
     *
     * @return 存在时返回键值对否则返回 {@code null}
     */
    public CustomProperties getCustomProperties() {
        ZipEntry entry = getEntry("docProps/custom.xml");
        if (entry == null) return null;
        Document document = null;
        // Load custom.xml
        SAXReader reader = SAXReaderUtil.createDefault();
        try {
            document = reader.read(zipFile.getInputStream(entry));
        } catch (DocumentException | IOException e) {
            LOGGER.warn("The file format is incorrect or corrupted. [docProps/custom.xml]");
        }
        if (document == null) return null;
        Element root = document.getRootElement();
        List<Element> list = root.elements();
        if (list == null || list.isEmpty()) return null;
        return CustomProperties.domToCustom(root);
    }
}

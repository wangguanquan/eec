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

package cn.ttzero.excel.reader;

import cn.ttzero.excel.annotation.NS;
import cn.ttzero.excel.annotation.TopNS;
import cn.ttzero.excel.entity.Relationship;
import cn.ttzero.excel.manager.Const;
import cn.ttzero.excel.manager.ExcelType;
import cn.ttzero.excel.manager.RelManager;
import cn.ttzero.excel.manager.docProps.App;
import cn.ttzero.excel.manager.docProps.Core;
import cn.ttzero.excel.util.FileUtil;
import cn.ttzero.excel.util.StringUtil;
import cn.ttzero.excel.util.ZipUtil;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.dom4j.*;
import org.dom4j.io.SAXReader;

import java.io.IOException;
import java.io.InputStream;
import java.io.UncheckedIOException;
import java.lang.reflect.Constructor;
import java.lang.reflect.Field;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.attribute.PosixFilePermissions;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.stream.Stream;
import java.util.stream.StreamSupport;

/**
 * Excel读取工具
 * 一个流式操作链，使用游标控制，游标只会向前，所以不能反复操作同一个Sheet流。
 * 同一个Sheet页内部Row对象是内存共享的，所以不要直接将Stream<Row>转为集合类.
 * 你首先应该考虑使用try-with-resource使用Reader或手动关闭ExcelReader。
 * <code>
 *     try (ExcelReader reader = ExcelReader.read(path)) {
 *         reader.sheets().flatMap(Sheet::rows).forEach(System.out::println);
 *     } catch (IOException e) {}
 * </code>
 * <p>
 * Create by guanquan.wang on 2018-09-22
 */
public class ExcelReader implements AutoCloseable {
    private Logger logger = LogManager.getLogger(getClass());

    protected ExcelReader() { }

    Path self;
    Sheet[] sheets;
    private Path temp;
    private ExcelType type;
    private AppInfo appInfo;


    /**
     * 实例化Reader
     *
     * @param path Excel路径
     * @return ExcelReader
     * @throws IOException 文件不存在或读取文件失败
     */
    public static ExcelReader read(Path path) throws IOException {
        return read(path, 0, 0);
    }

    /**
     * 实例化Reader
     *
     * @param stream Excel文件流
     * @return ExcelReader
     * @throws IOException 读取文件失败
     */
    public static ExcelReader read(InputStream stream) throws IOException {
        return read(stream, 0, 0);
    }

    /**
     * 实例化Reader
     *
     * @param path      Excel路径
     * @param cacheSize sharedString缓存大小，默认512
     *                  此参数影响读取文件次数
     * @return ExcelReader
     * @throws IOException 文件不存在或读取文件失败
     */
    public static ExcelReader read(Path path, int cacheSize) throws IOException {
        return read(path, cacheSize, 0);
    }

    /**
     * 实例化Reader
     *
     * @param stream    Excel文件流
     * @param cacheSize sharedString缓存大小，默认512
     *                  此参数影响读取文件次数
     * @return ExcelReader
     * @throws IOException 读取文件失败
     */
    public static ExcelReader read(InputStream stream, int cacheSize) throws IOException {
        return read(stream, cacheSize, 0);
    }

    /**
     * 实例化Reader
     *
     * @param path      Excel路径
     * @param cacheSize sharedString缓存大小，默认512
     *                  此参数影响读取文件次数
     * @param hotSize   热词区大小，默认512
     * @return ExcelReader
     * @throws IOException 文件不存在或读取文件失败
     */
    public static ExcelReader read(Path path, int cacheSize, int hotSize) throws IOException {
        return read(path, cacheSize, hotSize, false);
    }

    /**
     * 实例化Reader
     *
     * @param stream    Excel文件流
     * @param cacheSize sharedString缓存大小，默认512
     *                  将此参数影响读取文件次数
     * @param hotSize   热词区大小，默认512
     * @return ExcelReader
     * @throws IOException 读取文件失败
     */
    public static ExcelReader read(InputStream stream, int cacheSize, int hotSize) throws IOException {
        Path temp;
        if (FileUtil.isWindows()) {
            temp = Files.createTempFile(Const.EEC_PREFIX, null);
        } else {
            temp = Files.createTempFile(Const.EEC_PREFIX, null
                , PosixFilePermissions.asFileAttribute(PosixFilePermissions.fromString("rwxr-x---")));
        }
        if (temp == null) {
            throw new IOException("Create temp directory error. Please check your permission");
        }
        FileUtil.cp(stream, temp);
        return read(temp, cacheSize, hotSize, true);
    }

    /**
     * Type of excel
     *
     * @return enum type ExcelType
     */
    public ExcelType getType() {
        return type;
    }

    /**
     * to streams
     *
     * @return sheet流
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
     * get by index
     *
     * @param index sheet index of workbook
     * @return sheet
     */
    public Sheet sheet(int index) {
        try {
            return sheets[index].load(); // lazy loading worksheet data
        } catch (IOException e) {
            throw new UncheckedIOException(e);
        }
    }

    /**
     * get by name
     *
     * @param sheetName name
     * @return null if not found
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
     * get all sheets
     *
     * @return Sheet Array
     */
    public Sheet[] all() {
        return sheets;
    }

    /**
     * size of sheets
     *
     * @return int
     */
    public int getSize() {
        return sheets != null ? sheets.length : 0;
    }

    /**
     * close stream and delete temp files
     *
     * @throws IOException when fail close readers
     */
    public void close() throws IOException {
        // close sheet
        for (Sheet st : sheets) {
            st.close();
        }
        // delete temp files
        FileUtil.rm_rf(self.toFile(), true);
        if (temp != null) {
            FileUtil.rm(temp);
        }
    }

    /**
     * General information like title,subject and creator
     *
     * @return the information
     */
    public AppInfo getAppInfo() {
        if (appInfo == null) {
            appInfo = getGeneralInfo();
        }
        return appInfo;
    }


    // --- PRIVATE FUNCTIONS


    private ExcelReader(Path path, int cacheSize, int hotSize) throws IOException {
        // Store template stream as zip file
        Path temp = FileUtil.mktmp(Const.EEC_PREFIX);
        ZipUtil.unzip(Files.newInputStream(path), temp);

        // load workbook.xml
        SAXReader reader = new SAXReader();
        Document document;
        try {
            document = reader.read(Files.newInputStream(temp.resolve("xl/_rels/workbook.xml.rels")));
        } catch (DocumentException | IOException e) {
            FileUtil.rm_rf(temp.toFile(), true);
            throw new ExcelReadException(e);
        }
        @SuppressWarnings("unchecked")
        List<Element> list = document.getRootElement().elements();
        Relationship[] rels = new Relationship[list.size()];
        int i = 0;
        for (Element e : list) {
            rels[i++] = new Relationship(e.attributeValue("Id"), e.attributeValue("Target"), e.attributeValue("Type"));
        }
        RelManager relManager = RelManager.of(rels);

        try {
            document = reader.read(Files.newInputStream(temp.resolve("xl/workbook.xml")));
        } catch (DocumentException | IOException e) {
            // read style file fail.
            FileUtil.rm_rf(temp.toFile(), true);
            throw new ExcelReadException(e);
        }
        Element root = document.getRootElement();
        Namespace ns = root.getNamespaceForPrefix("r");

        // Load SharedString
        SharedString sst = new SharedString(temp.resolve("xl/sharedStrings.xml"), cacheSize, hotSize).load();

        List<Sheet> sheets = new ArrayList<>();
        @SuppressWarnings("unchecked")
        Iterator<Element> sheetIter = root.element("sheets").elementIterator();
        for (; sheetIter.hasNext(); ) {
            Element e = sheetIter.next();
            XMLSheet sheet = new XMLSheet();
            sheet.setName(e.attributeValue("name"));
            sheet.setIndex(Integer.parseInt(e.attributeValue("sheetId")));
            String state = e.attributeValue("state");
            sheet.setHidden("hidden".equals(state));
            Relationship r = relManager.getById(e.attributeValue(QName.get("id", ns)));
            if (r == null) {
                FileUtil.rm_rf(temp.toFile(), true);
                sheet.close();
                throw new ExcelReadException("File has be destroyed");
            }
            sheet.setPath(temp.resolve("xl").resolve(r.getTarget()));
            // put shared string
            sheet.setSst(sst);
            sheets.add(sheet);
        }

        // sort by sheet index
        sheets.sort(Comparator.comparingInt(Sheet::getIndex));

        Sheet[] sheets1 = new Sheet[sheets.size()];
        sheets.toArray(sheets1);

        this.sheets = sheets1;
        this.self = temp;
    }

    /**
     * 实例化Reader
     *
     * @param path      Excel路径
     * @param cacheSize sharedString缓存大小，默认512
     *                  此参数影响读取文件次数
     * @param hotSize   热词区大小，默认512
     * @param rmSource  是否删除源文件
     * @return ExcelReader
     * @throws IOException 文件不存在或读取文件失败
     */
    private static ExcelReader read(Path path, int cacheSize, int hotSize, boolean rmSource) throws IOException {
        // Check document type
        ExcelType type = getType(path);
        ExcelReader er;
        switch (type) {
            case XLSX:
                er = new ExcelReader(path, cacheSize, hotSize);
                break;
            case XLS:
                try {
                    Class<?> clazz = Class.forName("cn.ttzero.excel.reader.BIFF8Reader");
                    Constructor<?> constructor = clazz.getDeclaredConstructor(Path.class, int.class, int.class);
                    er = (ExcelReader) constructor.newInstance(path, cacheSize, hotSize);
                } catch (Exception e) {
                    throw new ExcelReadException("Only support read Office Open XML file.", e);
                }
                break;
            default:
                throw new ExcelReadException("Unknown file type.");
        }
        er.type = type;

        // storage source path
        if (rmSource) {
            er.temp = path;
        }

        return er;
    }

    /**
     * Check the documents type
     *
     * @param path documents path
     * @return enum of ExcelType
     */
    private static ExcelType getType(Path path) {
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
        int type = bytes[0] & 0xFF;
        type += (bytes[1] & 0xFF) << 8;
        type += (bytes[2] & 0xFF) << 16;
        type += (bytes[3] & 0xFF) << 24;

        int zip = 0x04034B50;
        int b1 = 0xE011CFD0;
        int b2 = 0xE11AB1A1;

        if (type == zip) {
            excelType = ExcelType.XLSX;
        } else if (type == b1 && length >= 8) {
            type = bytes[4] & 0xFF;
            type += (bytes[5] & 0xFF) << 8;
            type += (bytes[6] & 0xFF) << 16;
            type += (bytes[7] & 0xFF) << 24;
            if (type == b2) excelType = ExcelType.XLS;
        }
        return excelType;
    }

    protected AppInfo getGeneralInfo() {
        // load workbook.xml
        SAXReader reader = new SAXReader();
        Document document;
        try {
            document = reader.read(Files.newInputStream(self.resolve("docProps/app.xml")));
        } catch (DocumentException | IOException e) {
            throw new ExcelReadException(e);
        }
        Element root = document.getRootElement();
        App app = new App();
        app.setCompany(root.elementText("Company"));
        app.setApplication(root.elementText("Application"));
        app.setAppVersion(root.elementText("AppVersion"));
        app.setManager(root.elementText("Manager"));

        try {
            document = reader.read(Files.newInputStream(self.resolve("docProps/core.xml")));
        } catch (DocumentException | IOException e) {
            throw new ExcelReadException(e);
        }
        root = document.getRootElement();
        Core core = new Core();
        Class<Core> clazz = Core.class;
        TopNS topNS = clazz.getAnnotation(TopNS.class);
        String[] prefixs = topNS.prefix(), urls = topNS.uri();
        Field[] fields = clazz.getDeclaredFields();
        SimpleDateFormat format = new SimpleDateFormat("yyyy-MM-dd'T'hh:mm:ss'Z'");
        for (Field f : fields) {
            NS ns = f.getAnnotation(NS.class);
            if (ns == null) continue;

            f.setAccessible(true);
            int nsIndex = StringUtil.indexOf(prefixs, ns.value());
            if (nsIndex > -1) {
                Namespace namespace = new Namespace(ns.value(), urls[nsIndex]);
                Class<?> type = f.getType();
                String v = root.elementText(new QName(f.getName(), namespace));
                if (type == String.class) {
                    try {
                        f.set(core, v);
                    } catch (IllegalAccessException e) {
                        logger.warn("Set field (" + f + ") error.");
                    }
                } else if (type == Date.class) {
                    try {
                        f.set(core, format.parse(v));
                    } catch (ParseException | IllegalAccessException e) {
                        logger.warn("Set field (" + f + ") error.");
                    }
                }
            }
        }

        return new AppInfo(app, core);
    }
}

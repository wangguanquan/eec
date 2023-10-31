/*
 * Copyright (c) 2017-2018, guanquan.wang@yandex.com All Rights Reserved.
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
import org.ttzero.excel.util.DateUtil;
import org.ttzero.excel.util.FileUtil;
import org.ttzero.excel.util.StringUtil;

import java.io.Closeable;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
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

import static org.ttzero.excel.reader.SharedStrings.toInt;
import static org.ttzero.excel.util.FileUtil.exists;
import static org.ttzero.excel.util.StringUtil.isEmpty;
import static org.ttzero.excel.util.StringUtil.isNotEmpty;

/**
 * Excel Reader tools
 * <p>
 * A streaming operation chain, using cursor control, the cursor
 * will only move forward, so you cannot repeatedly operate the
 * same Sheet stream. If you need to read the data of a worksheet
 * multiple times please call the {@link Sheet#reset} method.
 * <p>
 * The internal Row object of the same Sheet is memory shared,
 * so don't directly convert Stream&lt;Row&gt; to a {@code Collection}.
 * You should first consider using the try-with-resource block to use Reader
 * or manually close the ExcelReader.
 * <blockquote><pre>
 * try (ExcelReader reader = ExcelReader.read(path)) {
 *     reader.sheets().flatMap(Sheet::rows).forEach(System.out::println);
 * } catch (IOException e) { }</pre></blockquote>
 *
 * @author guanquan.wang on 2018-09-22
 */
public class ExcelReader implements Closeable {
    private static final Logger LOGGER = LoggerFactory.getLogger(ExcelReader.class);

    /**
     * Specify {@link ExcelReader} only parse cell value (Default)
     */
    public static final int VALUE_ONLY = 0;
    /**
     * Parse cell value and calc
     */
    public static final int VALUE_AND_CALC = 1 << 1;
    /**
     * Copy value on merge cells
     */
    public static final int COPY_ON_MERGED = 1 << 2;

    protected ExcelReader() { }

    protected Sheet[] sheets;
    private Path temp;
    private ExcelType type;
    private AppInfo appInfo;
    /**
     * Store temporary files (Copy images to temp directory)
     */
    protected Path tempDir;
    /**
     * The Shared String Table
     */
    private SharedStrings sst;

    /**
     * Reader Option
     * <ul>
     * <li>0: only parse cell value (default)</li>
     * <li>2: parse cell value and calc</li>
     * <li>4: copy value on merge cells</li>
     * </ul>
     *
     * These attributes can be combined via `|`,
     * like: VALUE_ONLY|COPY_ON_MERGED
     */
    protected int option;

    /**
     * A formula flag
     */
    protected boolean hasFormula;

    /**
     * Picture or Tables
     */
    protected Drawings drawings;
    /**
     * A global styles
     */
    protected Styles styles;
    /**
     * Xlsx src file
     */
    protected ZipFile zipFile;

    /**
     * Constructor Excel Reader
     *
     * @param path the excel path
     * @return the {@link ExcelReader}
     * @throws IOException if path not exists or I/O error occur
     */
    public static ExcelReader read(Path path) throws IOException {
        return read(path, 0, 0, VALUE_ONLY);
    }

    /**
     * Constructor Excel Reader
     *
     * @param stream the {@link InputStream} of excel
     * @return the {@link ExcelReader}
     * @throws IOException if I/O error occur
     */
    public static ExcelReader read(InputStream stream) throws IOException {
        return read(stream, 0, 0, VALUE_ONLY);
    }

    /**
     * Constructor Excel Reader
     *
     * @param path the excel path
     * @param option the reader option.
     * @return the {@link ExcelReader}
     * @throws IOException if path not exists or I/O error occur
     */
    public static ExcelReader read(Path path, int option) throws IOException {
        return read(path, 0, 0, option);
    }

    /**
     * Constructor Excel Reader
     *
     * @param stream the {@link InputStream} of excel
     * @param option the reader option.
     * @return the {@link ExcelReader}
     * @throws IOException if I/O error occur
     */
    public static ExcelReader read(InputStream stream, int option) throws IOException {
        return read(stream, 0, 0, option);
    }

    /**
     * Constructor Excel Reader
     *
     * @param path       the excel path
     * @param bufferSize the {@link SharedStrings} buffer size. default is 512
     *                   This parameter affects the number of read times.
     * @param option the reader option.
     * @return the {@link ExcelReader}
     * @throws IOException if path not exists or I/O error occur
     */
    public static ExcelReader read(Path path, int bufferSize, int option) throws IOException {
        return read(path, bufferSize, 0, option);
    }

    /**
     * Constructor Excel Reader
     *
     * @param stream     the {@link InputStream} of excel
     * @param bufferSize the {@link SharedStrings} buffer size. default is 512
     *                   This parameter affects the number of read times.
     * @param option the reader option.
     * @return the {@link ExcelReader}
     * @throws IOException if I/O error occur
     */
    public static ExcelReader read(InputStream stream, int bufferSize, int option) throws IOException {
        return read(stream, bufferSize, 0, option);
    }

    /**
     * Constructor Excel Reader
     *
     * @param stream     the {@link InputStream} of excel
     * @param bufferSize the {@link SharedStrings} buffer size. default is 512
     *                   This parameter affects the number of read times.
     * @param cacheSize  the {@link Cache} size, default is 512
     * @param option the reader option.
     * @return the {@link ExcelReader}
     * @throws IOException if I/O error occur
     */
    public static ExcelReader read(InputStream stream, int bufferSize, int cacheSize, int option) throws IOException {
        Path temp = FileUtil.mktmp(Const.EEC_PREFIX);
        if (temp == null) {
            throw new IOException("Create temp directory error. Please check your permission");
        }
        FileUtil.cp(stream, temp);
        ExcelReader reader = read(temp, bufferSize, cacheSize, option);
        reader.temp = temp;
        return reader;
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
     * @return {@link Stream} of {@link Sheet}
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
     * Returns all sheets
     *
     * @return Sheet Array
     */
    public Sheet[] all() {
        return sheets;
    }

    /**
     * Size of sheets
     *
     * @return int
     */
    public int getSize() {
        return sheets != null ? sheets.length : 0;
    }

    /**
     * Close stream and delete temp files
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
        if (sst != null) sst.close();

        // Close source file
        if (zipFile != null) zipFile.close();

        // Remove temp file
        if (temp != null) FileUtil.rm(temp);

        // Remove temp dir
        if (tempDir != null) FileUtil.rm_rf(tempDir);
    }

    /**
     * General information like title,subject and creator
     *
     * @return the information
     */
    public AppInfo getAppInfo() {
        return appInfo != null ? appInfo : (appInfo = getGeneralInfo());
    }

    /**
     * Check current workbook has formula
     *
     * @return boolean
     * @deprecated This method can not accurately reflect whether
     * the workbook contains formulas. Here we only check whether
     * they contain a common calcChain file, and the Excel files
     * generated by some tools do not contain the calcChain file,
     * but write all formulas in each cell.
     */
    @Deprecated
    public boolean hasFormula() {
        return this.hasFormula;
    }

    /**
     * Make the reader parse formula
     *
     * @return {@link ExcelReader}
     */
    public ExcelReader parseFormula() {
        if (hasFormula) {
            ZipEntry entry = getEntry("xl/calcChain.xml");
            if (entry == null) return this;
            long[][] calcArray = null;
            try (InputStream is = zipFile.getInputStream(entry)) {
                // Formula string if exists
                calcArray = parseCalcChain(is);
            } catch (IOException ex) {
                LOGGER.warn("Part of `calcChain` has be damaged, It will be ignore all formulas.", ex);
            }

            if (calcArray == null) return this;
            int i = 0;
            for (int n; i < sheets.length; i++) {
                n = sheets[i].getId();
                if (calcArray[n - 1] == null) continue;
                if (!(sheets[i] instanceof CalcSheet)) {
                    sheets[i] = sheets[i].asCalcSheet();
                }
                if (sheets[i] instanceof XMLCalcSheet) {
                    ((XMLCalcSheet) sheets[i]).setCalc(calcArray[n - 1]);
                }
            }
        } else {
            for (Sheet sheet : sheets) {
                sheet.asCalcSheet();
            }
        }
        return this;
    }

    /**
     * Copy values when reading merged cells.
     * <p>
     * By default, the values of the merged cells are only
     * stored in the first Cell, and other cells have no values.
     * Call this method to copy the value to other cells in the merge.
     * <blockquote><pre>
     * |---------|     |---------|     |---------|
     * |         |     |  1 |    |     |  1 |  1 |
     * |    1    |  =&gt; |----|----|  =&gt; |----|----|
     * |         |     |    |    |     |  1 |  1 |
     * |---------|     |---------|     |---------|
     * Merged(A1:B2)     Default           Copy
     *                  Value in A1
     *                  others are
     *                  `null`
     * </pre></blockquote>
     *
     * @return {@link ExcelReader}
     */
    public ExcelReader copyOnMergeCells() {
        for (int i = 0; i < sheets.length; i++) {
            if (sheets[i] instanceof MergeSheet) continue;
            sheets[i] = sheets[i].asMergeSheet();
        }
        return this;
    }

    // --- PROTECTED FUNCTIONS

    public static final Set<String> MUST_CHECK_PART = new HashSet<>(Arrays.asList(Const.ContentType.WORKBOOK
            , Const.ContentType.SHAREDSTRING, Const.ContentType.SHEET, Const.ContentType.STYLE));

    protected ContentType checkContentType() {
        // Read [Content_Types].xml
        ZipEntry entry = getEntry("[Content_Types].xml");
        if (entry == null) {
            if (temp != null) FileUtil.rm(temp);
            throw new ExcelReadException("The file format is incorrect or corrupted. [[Content_Types].xml]");
        }
        SAXReader reader = SAXReader.createDefault();
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

    public ExcelReader(InputStream is) throws IOException {
        this(is, 0, 0, VALUE_ONLY);
    }

    public ExcelReader(InputStream is, int option) throws IOException {
        this(is, 0, 0, option);
    }

    public ExcelReader(InputStream stream, int bufferSize, int cacheSize, int option) throws IOException {
        Path temp = FileUtil.mktmp(Const.EEC_PREFIX);
        if (temp == null) {
            throw new IOException("Create temp directory error. Please check permission");
        }
        FileUtil.cp(stream, temp);

        init(temp, bufferSize, cacheSize, option);
    }

    public ExcelReader(Path path) throws IOException {
        init(path, 0, 0, VALUE_ONLY);
    }

    public ExcelReader(Path path, int option) throws IOException {
        init(path, 0, 0, option);
    }

    public ExcelReader(Path path, int bufferSize, int cacheSize, int option) throws IOException {
        init(path, bufferSize, cacheSize, option);
    }

    protected ExcelReader init(Path path, int bufferSize, int cacheSize, int option) throws IOException {
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
        SAXReader reader = SAXReader.createDefault();
        Document document;

        // Load SharedString
        ZipEntry entry = getEntry("xl/sharedStrings.xml");
        if (entry != null) {
            sst = new SharedStrings(zipFile.getInputStream(entry), bufferSize, cacheSize).load();
        }

        // Load Styles
        entry = getEntry("xl/styles.xml");
        if (entry != null) {
            try {
                styles = Styles.load(zipFile.getInputStream(entry));
                // Load Theme style
                ZipEntry themeEntry = getEntry("theme/theme1.xml");
                if (themeEntry != null) Theme.load(zipFile.getInputStream(themeEntry));
            } catch (Exception ex) {
                LOGGER.warn("Parse style failed.", ex);
            }
        }
        // Construct a empty Styles
        if (styles == null) {
            styles = Styles.forReader();
        }

        this.option = option;
        hasFormula = getEntry("xl/calcChain.xml") != null;

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
        for (; sheetIter.hasNext(); ) {
            Element e = sheetIter.next();
            XMLSheet sheet = (XMLSheet) sheetFactory(option);
            sheet.setName(e.attributeValue("name"));
            sheet.setId(Integer.parseInt(e.attributeValue("sheetId")));
            String state = e.attributeValue("state");
            sheet.setHidden("hidden".equals(state));
            Relationship r = relManager.getById(e.attributeValue(QName.get("id", ns)));
            if (r == null) {
                sheet.close();
                throw new ExcelReadException("The file format is incorrect or corrupted.");
            }
            sheet.setPath("xl/" + r.getTarget());
            entry = getEntry("xl/" + r.getTarget());
            if (entry == null) {
                sheet.close();
                throw new ExcelReadException("The file format is incorrect or corrupted.");
            }
            sheet.setZipFile(zipFile);
            sheet.setZipEntry(entry);
            // put shared string
            sheet.setSst(sst);
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

        if ((option >> 1 & 1) == 1) {
            parseFormula();
        }
        return this;
    }

    /**
     * Constructor Excel Reader
     *
     * @param path       the excel path
     * @param bufferSize the {@link SharedStrings} buffer size. default is 512
     *                   This parameter affects the number of read times.
     * @param cacheSize  the {@link Cache} size, default is 512
     * @param option the reader option.
     * @return the {@link ExcelReader}
     * @throws FileNotFoundException if the path not exists or no permission to read
     * @throws IOException if I/O error occur
     */
    public static ExcelReader read(Path path, int bufferSize, int cacheSize, int option) throws IOException {
        if (!exists(path)) {
            throw new FileNotFoundException(path.toString());
        }
        // Check document type
        ExcelType type = getType(path);
        LOGGER.debug("File type: {}", type);
        ExcelReader er;
        switch (type) {
            case XLSX:
                er = new ExcelReader(path, bufferSize, cacheSize, option);
                break;
            case XLS:
                try {
                    Class<?> clazz = Class.forName("org.ttzero.excel.reader.BIFF8Reader");
                    Constructor<?> constructor = clazz.getDeclaredConstructor(Path.class, int.class, int.class, int.class);
                    er = (ExcelReader) constructor.newInstance(path, bufferSize, cacheSize, option);
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
        er.type = type;

        return er;
    }

    /**
     * Create a read sheet
     *
     * @param option the reader option.
     * @return Sheet extends XMLSheet
     */
    protected Sheet sheetFactory(int option) {
        XMLSheet sheet;
        switch (option) {
            case VALUE_AND_CALC: sheet = new XMLSheet().asCalcSheet(); break;
            case COPY_ON_MERGED: sheet = new XMLSheet().asMergeSheet(); break;
            // TODO full reader
//            case VALUE_AND_CALC|COPY_ON_MERGED: break;
            default            : sheet = new XMLSheet();
        }
        return sheet;
    }

    /**
     * Check the documents type
     *
     * @param path documents path
     * @return enum of ExcelType
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

    protected AppInfo getGeneralInfo() {
        // load app.xml
        SAXReader reader = SAXReader.createDefault();
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

    /* Parse `calcChain` */
    static long[][] parseCalcChain(InputStream is) {
        SAXReader reader = SAXReader.createDefault();
        Element calcChain;
        try {
            calcChain = reader.read(is).getRootElement();
        } catch (DocumentException e) {
            LOGGER.warn("Part of `calcChain` has be damaged, It will be ignore all formulas.");
            return null;
        }

        Iterator<Element> ite = calcChain.elementIterator();
        int i = 1, n = 10;
        long[][] array = new long[n][];
        int[] indices = new int[n];
        for (; ite.hasNext(); ) {
            Element e = ite.next();
            String si = e.attributeValue("i"), r = e.attributeValue("r");
            if (isNotEmpty(si)) {
                i = toInt(si.toCharArray(), 0, si.length());
            }
            if (isNotEmpty(r)) {
                if (n < i) {
                    n <<= 1;
                    indices = Arrays.copyOf(indices, n);
                    long[][] _array = new long[n][];
                    for (int j = 0; j < n; j++) _array[j] = array[j]; // Do not copy hear.
                    array = _array;
                }
                long[] sub = array[i - 1];
                if (sub == null) {
                    sub = new long[10];
                    array[i - 1] = sub;
                }

                if (++indices[i - 1] > sub.length) {
                    long[] _sub = new long[sub.length << 1];
                    System.arraycopy(sub, 0, _sub, 0, sub.length);
                    array[i - 1] = sub = _sub;
                }
                sub[indices[i - 1] - 1] = cellRangeToLong(r);
            }
        }

        i = 0;
        for (; i < n; i++) {
            if (indices[i] > 0) {
                long[] a = Arrays.copyOf(array[i], indices[i]);
                Arrays.sort(a);
                array[i] = a;
            } else array[i] = null;
        }
        return array;
    }

    /**
     * Cell range string convert to long
     * 0-16: column number
     * 17-48: row number
     * <blockquote><pre>
     * range string| long value
     * ------------|------------
     * A1          | 65537
     * AA10        | 655387
     * </pre></blockquote>
     * @param r the range string of cell
     * @return long value
     */
    public static long cellRangeToLong(String r) {
        char[] values = r.toCharArray();
        long v = 0L;
        int n = 0;
        for (char value : values) {
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
     * List all pictures in excel
     *
     * @return picture list or null if not exists.
     */
    public List<Drawings.Picture> listPictures() {
        return drawings != null ? drawings.listPictures() : null;
    }

    /**
     * Returns a global {@link Styles}
     *
     * @return a style entry
     */
    public Styles getStyles() {
        return styles;
    }

    /**
     * 读取Zip包中的实体
     *
     * @param name 实体名
     * @return 如果实体存在则返回 {@link ZipEntry} 否则返回{@code null}
     */
    public ZipEntry getEntry(String name) {
        return getEntry(zipFile, name);
    }

    /**
     * Find {@code ZipEntry} by name
     *
     * @param zipFile src zip file
     * @param name entry name
     * @return {@code ZipEntry} if exist, otherwise null
     */
    public static ZipEntry getEntry(ZipFile zipFile, String name) {
        char c0 = name.charAt(0);
        if (c0 == '/' || c0 == '\\') name = name.substring(1);
        ZipEntry entry = zipFile.getEntry(name);
        if (entry == null) entry = zipFile.getEntry(name.replace('/', '\\'));
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
     * Simple convert to zip path
     *
     * @param path file path
     * @return zip path
     */
    public static String toZipPath(String path) {
        int i = 0;
        if (path.startsWith("../") || path.startsWith("..\\")) i = 3;
        else if (path.startsWith("./") || path.startsWith(".\\")) i = 2;
        else if (path.charAt(0) == '/' || path.charAt(0) == '\\') i = 1;
        return i > 0 ? path.substring(i) : path;
    }
}

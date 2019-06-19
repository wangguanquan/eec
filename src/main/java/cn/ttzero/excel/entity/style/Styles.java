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

package cn.ttzero.excel.entity.style;

import cn.ttzero.excel.annotation.TopNS;
import cn.ttzero.excel.entity.I18N;
import cn.ttzero.excel.entity.Storageable;
import cn.ttzero.excel.manager.Const;
import cn.ttzero.excel.reader.ExcelReadException;
import cn.ttzero.excel.util.FileUtil;
import org.dom4j.Document;
import org.dom4j.DocumentException;
import org.dom4j.DocumentFactory;
import org.dom4j.Element;
import org.dom4j.io.SAXReader;

import java.awt.Color;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.*;

import static cn.ttzero.excel.util.StringUtil.isEmpty;
import static cn.ttzero.excel.util.StringUtil.isNotEmpty;

/**
 * Each excel style consists of {@link NumFmt}, {@link Font},
 * {@link Fill}, {@link Border}, {@link Verticals} and
 * {@link Horizontals}, each Worksheet introduced by subscript.
 * <p>
 * EEC uses an Integer value to represent a complete pattern,
 * 0-8 bits are stored in NumFmt, 8-14 bits are stored in Font,
 * 14-20 bits are stored in Fill, 20-26 bits are stored in Border,
 * 26-29 bits are stored in Vertical and 29-32 bits are stored in
 * Horizontal. The Build-In number format does not write into styles.
 * <p>
 * Created by guanquan.wang on 2017/10/13.
 */
@TopNS(prefix = "", uri = Const.SCHEMA_MAIN, value = "styleSheet")
public class Styles implements Storageable {

    private Map<Integer, Integer> map;
    private Document document;

    private List<Font> fonts;
    private List<NumFmt> numFmts;
    private List<Fill> fills;
    private List<Border> borders;

    /**
     * Cache the data/time format style index.
     * It's use for fast test the cell value is a data or time value
     */
    private Set<Integer> dateFmtCache;

    private Styles() {
        map = new HashMap<>();
    }

    /**
     * Returns the style index, if the style not exists it will
     * be insert into styles
     *
     * @param s the value of style
     * @return the style index
     */
    public int of(int s) {
        int n = map.getOrDefault(s, 0);
        if (n == 0) {
            n = addStyle(s);
            map.put(s, n);
        }
        return n;
    }

    /**
     * Returns the number of styles
     *
     * @return the total styles
     */
    public int size() {
        return map.size();
    }

    static final int INDEX_NUMBER_FORMAT = 24;
    static final int INDEX_FONT = 18;
    static final int INDEX_FILL = 12;
    static final int INDEX_BORDER = 6;
    static final int INDEX_VERTICAL = 3;
    static final int INDEX_HORIZONTAL = 0;

    /**
     * Create a general style
     *
     * @param i18N the {@link I18N}
     * @return Styles
     */
    public static Styles create(I18N i18N) {
        Styles self = new Styles();

        DocumentFactory factory = DocumentFactory.getInstance();
        TopNS ns = self.getClass().getAnnotation(TopNS.class);
        Element rootElement;
        if (ns != null) {
            rootElement = factory.createElement(ns.value(), ns.uri()[0]);
        } else {
            rootElement = factory.createElement("styleSheet", Const.SCHEMA_MAIN);
        }
        // number format
        rootElement.addElement("numFmts").addAttribute("count", "0");
        // font
        rootElement.addElement("fonts").addAttribute("count", "0");
        // fill
        rootElement.addElement("fills").addAttribute("count", "0");
        // border
        rootElement.addElement("borders").addAttribute("count", "0");
        // cellStyleXfs
        Element cellStyleXfs = rootElement.addElement("cellStyleXfs").addAttribute("count", "1");
        cellStyleXfs.addElement("xf")   // General style
            .addAttribute("borderId", "0")
            .addAttribute("fillId", "0")
            .addAttribute("fontId", "0")
            .addAttribute("numFmtId", "0")
            .addElement("alignment")
            .addAttribute("vertical", "center");
        // cellXfs
        rootElement.addElement("cellXfs").addAttribute("count", "0");
        // cellStyles
        Element cellStyles = rootElement.addElement("cellStyles").addAttribute("count", "1");
        cellStyles.addElement("cellStyle")
            .addAttribute("builtinId", "0")
            .addAttribute("name", i18N.get("general"))
            .addAttribute("xfId", "0");

        self.document = factory.createDocument(rootElement);

        self.numFmts = new ArrayList<>();
        self.addNumFmt(new NumFmt("yyyy\\-mm\\-dd"));
        self.addNumFmt(new NumFmt("yyyy\\-mm\\-dd\\ hh:mm:ss"));

        self.fonts = new ArrayList<>();
        Font font1 = new Font(i18N.get("en-font-family"), 11, Color.black);  // en
        font1.setFamily(2);
        font1.setScheme("minor");
        self.addFont(font1);

        String lang = Locale.getDefault().toLanguageTag();
        // 添加中文默认字体
        Font font2 = new Font(i18N.get("local-font-family"), 11); // cn
        font2.setFamily(3);
        font2.setScheme("minor");
        if ("zh-CN".equals(lang)) {
            font2.setCharset(Charset.GB2312);
        } else if ("zh-TW".equals(lang)) {
            font2.setCharset(Charset.CHINESEBIG5);
        }
        // TODO other charset
        self.addFont(font2);

        self.fills = new ArrayList<>();
        self.addFill(new Fill(PatternType.none));
        self.addFill(new Fill(PatternType.gray125));

        self.borders = new ArrayList<>();
        self.addBorder(Border.parse("none"));
        self.addBorder(Border.parse("thin black"));

        // cellXfs
        self.of(0); // General

        return self;
    }

    /**
     * Load the style file from disk
     *
     * @param path the style file path
     * @return the {@link Styles} Object
     */
    @SuppressWarnings("unchecked")
    public static Styles load(Path path) {
        // load workbook.xml
        SAXReader reader = new SAXReader();
        Document document;
        try {
            document = reader.read(Files.newInputStream(path));
        } catch (DocumentException | IOException e) {
            throw new ExcelReadException(e);
        }

        Styles self = new Styles();
        Element root = document.getRootElement();
        // Number format
        Element numFmts = root.element("numFmts");
        // Break if there don't contains 'numFmts' tag
        if (numFmts == null) {
            return self;
        }
        List<Element> sub = numFmts.elements();
        self.numFmts = new ArrayList<>();
        for (Element e : sub) {
            String id = getAttr(e, "numFmtId"), code = getAttr(e, "formatCode");
            self.numFmts.add(new NumFmt(Integer.parseInt(id), code));
        }
        // Sort by id
        self.numFmts.sort(Comparator.comparingInt(NumFmt::getId));

        /*
        Ignore other styles, there only parse number format
        It's use for Excel reader
         */

        // Cell xf
        Element cellXfs = root.element("cellXfs");
        sub = cellXfs.elements();
        int i = 0;
        for (Element e : sub) {
            String applyNumberFormat = getAttr(e, "applyNumberFormat");
            if (isNotEmpty(applyNumberFormat) && Integer.parseInt(applyNumberFormat) == 1) {
                String numFmtId = getAttr(e, "numFmtId");
                int style = Integer.parseInt(numFmtId) << INDEX_NUMBER_FORMAT;
                self.map.put(i, style);
            }
            i++;
        }
        // Test number format
        for (Integer styleIndex : self.map.keySet()) {
            self.isDate(styleIndex);
        }

        return self;
    }

    /**
     * Add number format
     *
     * @param numFmt the {@link NumFmt} entry
     * @return the numFmt part value in style
     */
    public final int addNumFmt(NumFmt numFmt) {
        // check and search default code
        if (numFmt.getId() < 0) {
            if (isEmpty(numFmt.getCode())) {
                throw new NullPointerException("NumFmt code");
            }
            int index = BuiltInNumFmt.indexOf(numFmt.getCode());
            if (index > -1) { // default code
                numFmt.setId(index);
            } else {
                int i = numFmts.indexOf(numFmt);
                if (i <= -1) {
                    int id;
                    if (numFmts.isEmpty()) {
                        id = 176; // customer id
                    } else {
                        id = numFmts.get(numFmts.size() - 1).getId() + 1;
                    }
                    numFmt.setId(id);
                    numFmts.add(numFmt);

                    Element element = document.getRootElement().element("numFmts");
                    element.attribute("count").setValue(String.valueOf(numFmts.size()));
                    numFmt.toDom4j(element);
                } else {
                    numFmt.setId(numFmts.get(i).getId());
                }
            }
        }
        return numFmt.getId() << INDEX_NUMBER_FORMAT;
    }

    /**
     * Add font
     *
     * @param font the {@link Font} entry
     * @return the font part value in style
     */
    public final int addFont(Font font) {
        if (isEmpty(font.getName())) {
            throw new FontParseException("Font name not support.");
        }
        int i = fonts.indexOf(font);
        if (i <= -1) {
            fonts.add(font);
            i = fonts.size() - 1;

            Element element = document.getRootElement().element("fonts");
            element.attribute("count").setValue(String.valueOf(fonts.size()));
            font.toDom4j(element);
        }
        return i << INDEX_FONT;
    }

    /**
     * Add fill
     *
     * @param fill the {@link Fill} entry
     * @return the fill part value in style
     */
    public final int addFill(Fill fill) {
        int i = fills.indexOf(fill);
        if (i <= -1) {
            fills.add(fill);
            i = fills.size() - 1;
            Element element = document.getRootElement().element("fills");
            element.attribute("count").setValue(String.valueOf(fills.size()));
            fill.toDom4j(element);
        }
        return i << INDEX_FILL;
    }

    /**
     * Add border
     *
     * @param border the {@link Border} entry
     * @return the border part value in style
     */
    public final int addBorder(Border border) {
        int i = borders.indexOf(border);
        if (i <= -1) {
            borders.add(border);
            i = borders.size() - 1;
            Element element = document.getRootElement().element("borders");
            element.attribute("count").setValue(String.valueOf(borders.size()));
            border.toDom4j(element);
        }
        return i << INDEX_BORDER;
    }

    private static int[] unpack(int style) {
        int[] styles = new int[6];
        styles[0] = style >>> INDEX_NUMBER_FORMAT;
        styles[1] = style << 8 >>> (INDEX_FONT + 8);
        styles[2] = style << 14 >>> (INDEX_FILL + 14);
        styles[3] = style << 20 >>> (INDEX_BORDER + 20);
        styles[4] = style << 26 >>> (INDEX_VERTICAL + 26);
        styles[5] = style << 29 >>> (INDEX_HORIZONTAL + 29);
        return styles;
    }

    private static int pack(int[] styles) {
        return styles[0] << INDEX_NUMBER_FORMAT
            | styles[1] << INDEX_FONT
            | styles[2] << INDEX_FILL
            | styles[3] << INDEX_BORDER
            | styles[4] << INDEX_VERTICAL
            | styles[5] << INDEX_HORIZONTAL
            ;
    }

    static final String[] attrNames = {"numFmtId", "fontId", "fillId", "borderId", "vertical", "horizontal"
        , "applyNumberFormat", "applyFont", "applyFill", "applyBorder", "applyAlignment"};

    /**
     * add style in document
     *
     * @param s style
     * @return style index in styles array.
     */
    private synchronized int addStyle(int s) {
        int[] styles = unpack(s);
        Element root = document.getRootElement();
        Element cellXfs = root.element("cellXfs");
        int count;
        if (cellXfs == null) {
            cellXfs = root.addElement("cellXfs").addAttribute("count", "0");
            count = 0;
        } else {
            count = Integer.parseInt(cellXfs.attributeValue("count"));
        }

        int n = cellXfs.elements().size();
        Element newXf = cellXfs.addElement("xf");
        newXf.addAttribute(attrNames[0], String.valueOf(styles[0]))
            .addAttribute(attrNames[1], String.valueOf(styles[1]))
            .addAttribute(attrNames[2], String.valueOf(styles[2]))
            .addAttribute(attrNames[3], String.valueOf(styles[3]))
            .addAttribute("xfId", "0")
        ;
        int start = 6;
        if (styles[0] > 0) {
            newXf.addAttribute(attrNames[start], "1");
        }
        if (styles[1] > 0) {
            newXf.addAttribute(attrNames[start + 1], "1");
        }
        if (styles[2] > 0) {
            newXf.addAttribute(attrNames[start + 2], "1");
        }
        if (styles[3] > 0) {
            newXf.addAttribute(attrNames[start + 3], "1");
        }
        if ((styles[4] | styles[5]) > 0) {
            newXf.addAttribute(attrNames[start + 4], "1");
        }

        Element subEle = newXf.addElement("alignment").addAttribute(attrNames[4], Verticals.of(styles[4]));
        if (styles[5] > 0) {
            subEle.addAttribute(attrNames[5], Horizontals.of(styles[5]));
        }
        cellXfs.addAttribute("count", String.valueOf(count + 1));
        return n;
    }

    /**
     * Write style to disk
     *
     * @param styleFile the storage path
     * @throws IOException if I/O error occur
     */
    @Override
    public void writeTo(Path styleFile) throws IOException {
        if (document != null) { // Not null
            FileUtil.writeToDiskNoFormat(document, styleFile);
        } else {
            Files.copy(getClass().getClassLoader().getResourceAsStream("template/styles.xml"), styleFile);
        }
    }

    int indexOf(int[] array, int v) {
        for (int i = 0; i < array.length; i++) {
            if (array[i] == v) {
                return i;
            }
        }
        return -1;
    }

    ////////////////////////clear style///////////////////////////////
    public static int clearNumfmt(int style) {
        return style & (-1 >>> 32 - INDEX_NUMBER_FORMAT);
    }

    public static int clearFont(int style) {
        return style & ~((-1 >>> 32 - (INDEX_NUMBER_FORMAT - INDEX_FONT)) << INDEX_FONT);
    }

    public static int clearFill(int style) {
        return style & ~((-1 >>> 32 - (INDEX_FONT - INDEX_FILL)) << INDEX_FILL);
    }

    public static int clearBorder(int style) {
        return style & ~((-1 >>> 32 - (INDEX_FILL - INDEX_BORDER)) << INDEX_BORDER);
    }

    public static int clearVertical(int style) {
        return style & ~((-1 >>> 32 - (INDEX_BORDER - INDEX_VERTICAL)) << INDEX_VERTICAL);
    }

    public static int clearHorizontal(int style) {
        return style & ~(-1 >>> 32 - (INDEX_VERTICAL - INDEX_HORIZONTAL));
    }

    ////////////////////////reset style/////////////////////////////
    public static int reset(int style, int newStyle) {
        int[] sub = unpack(style), nsub = unpack(newStyle);
        for (int i = 0; i < sub.length; i++) {
            if (nsub[i] > 0) {
                sub[i] = nsub[i];
            }
        }
        return pack(sub);
    }

    ////////////////////////default border style/////////////////////////////
    public static int defaultCharBorderStyle() {
        return (1 << INDEX_BORDER) | Horizontals.CENTER_CONTINUOUS;
    }

    public static int defaultStringBorderStyle() {
        return (1 << INDEX_FONT) | (1 << INDEX_BORDER) | Horizontals.LEFT;
    }

    public static int defaultIntBorderStyle() {
        return (1 << INDEX_NUMBER_FORMAT) | (1 << INDEX_BORDER) | Horizontals.RIGHT;
    }

    public static int defaultDateBorderStyle() {
        return (176 << INDEX_NUMBER_FORMAT) | (1 << INDEX_BORDER) | Horizontals.CENTER;
    }

    public static int defaultTimestampBorderStyle() {
        return (177 << INDEX_NUMBER_FORMAT) | (1 << INDEX_BORDER) | Horizontals.CENTER;
    }

    public static int defaultDoubleBorderStyle() {
        return (2 << INDEX_NUMBER_FORMAT) | (1 << INDEX_FONT) | (1 << INDEX_BORDER) | Horizontals.RIGHT;
    }

    ////////////////////////default style/////////////////////////////
    public static int defaultCharStyle() {
        return Horizontals.CENTER_CONTINUOUS;
    }

    public static int defaultStringStyle() {
        return (1 << INDEX_FONT) | Horizontals.LEFT;
    }

    public static int defaultIntStyle() {
        return (1 << INDEX_NUMBER_FORMAT) | Horizontals.RIGHT;
    }

    public static int defaultDateStyle() {
        return (176 << INDEX_NUMBER_FORMAT) | Horizontals.CENTER;
    }

    public static int defaultTimestampStyle() {
        return (177 << INDEX_NUMBER_FORMAT) | Horizontals.CENTER;
    }

    public static int defaultDoubleStyle() {
        return (2 << INDEX_NUMBER_FORMAT) | (1 << INDEX_FONT) | Horizontals.RIGHT;
    }

    ////////////////////////////////Check style////////////////////////////////
    public static boolean hasNumFmt(int style) {
        return style >>> INDEX_NUMBER_FORMAT != 0;
    }

    public static boolean hasFont(int style) {
        return style << 8 >>> (INDEX_FONT + 8) != 0;
    }

    public static boolean hasFill(int style) {
        return style << 14 >>> (INDEX_FILL + 14) != 0;
    }

    public static boolean hasBorder(int style) {
        return style << 20 >>> (INDEX_BORDER + 20) != 0;
    }

    public static boolean hasVertical(int style) {
        return style << 26 >>> (INDEX_VERTICAL + 26) != 0;
    }

    public static boolean hasHorizontal(int style) {
        return style << 29 >>> (INDEX_HORIZONTAL + 29) != 0;
    }

    ////////////////////////////////To object//////////////////////////////////
    public NumFmt getNumFmt(int style) {
        return numFmts.get(style >>> INDEX_NUMBER_FORMAT);
    }

    public Fill getFill(int style) {
        return fills.get(style << 14 >>> (INDEX_FILL + 14));
    }

    public Font getFont(int style) {
        return fonts.get(style << 8 >>> (INDEX_FONT + 8));
    }

    public Border getBorder(int style) {
        return borders.get(style << 20 >>> (INDEX_BORDER + 20));
    }

    public int getVertical(int style) {
        return style << 26 >>> (INDEX_VERTICAL + 26);
    }

    /**
     * Returns the attribute value from Element
     *
     * @param element current element
     * @param attr    the attr name
     * @return the attr value
     */
    public static String getAttr(Element element, String attr) {
        return element != null ? element.attributeValue(attr) : null;
    }

    /**
     * Test the style is data format
     *
     * @param styleIndex the style index
     * @return true if the style content data format
     */
    public boolean isDate(int styleIndex) {
        // Test from cache
        if (fastTestDateFmt(styleIndex)) return true;

        Integer style = map.get(styleIndex);
        if (style == null) return false;
        int nf = style >> INDEX_NUMBER_FORMAT & 0xFF;

        boolean isDate = false;
        NumFmt numFmt;
        // All indexes from 0 to 163 are reserved for built-in formats.
        // The first user-defined format starts at 164.
        if (nf < 164) {
            isDate = (nf >= 14 && nf <= 22
                || nf >= 27 && nf <= 36
                || nf >= 45 && nf <= 47
                || nf >= 50 && nf <= 58
                || nf == 81);
            // Test by numFmt code
        } else if ((numFmt = findFmtById(nf)) != null && (isNotEmpty(numFmt.getCode()))) {
            isDate = testCodeIsDate(numFmt.getCode());
        }

        // Put into data/time format cache
        // Ignore the style code, Uniform use of 'yyyy-mm-dd hh:mm:ss' format output
        if (isDate) {
            if (dateFmtCache == null) dateFmtCache = new HashSet<>();
            dateFmtCache.add(styleIndex);
        }
        return isDate;
    }

    public static boolean testCodeIsDate(String code) {
        char[] chars = code.toCharArray();

        int score = 0;
        byte[] byteScore = new byte[26];
        for (int i = 0, size = chars.length; i < size; ) {
            char c = chars[i];
            // To lower case
            if (c >= 65 && c <= 90) c += 32;

            int a = ++i;

            if (c == '[') {
                // Found the end char ']'
                for (; i < size && chars[i] != ']'; i++) ;
                int len = i - a + 1;
                // DBNum{n}
                if (len == 6 && chars[a] == 'D' && chars[a + 1] == 'B' && chars[a + 2] == 'N'
                    && chars[a + 3] == 'u' && chars[a + 4] == 'm') {
                    int n = chars[a + 5] - '0';
                    // If use "[DBNum{n}]" etc. as the Excel display format, you can use Chinese numerals.
                    if (n != 1 && n != 2 && n != 3) break;
                }
                // Maybe is a LCID
                // [$-xxyyzzzz]
                // https://stackoverflow.com/questions/54134729/what-does-the-130000-in-excel-locale-code-130000-mean
                else if (i - a > 2 && chars[a] == '$' && chars[a + 1] == '-') {
                    // Language & Calendar Identifier
//                    String lcid = new String(chars, a + 2, i - a - 2);
                    // Maybe the format is a data
                    score = 50;
                }
            } else {
                switch (c) {
                    case 'y':
                    case 'm':
                    case 'd':
                    case 'h':
                    case 's':
                        for (; i < size && chars[i] == c; i++) ;
                        byteScore[c - 'a'] += i - a + 1;
                        break;
                    case 'a':
                    case 'p':
                        if (a < size && chars[a] == 'm')
                            byteScore[c - 'a'] += 1;
                        break;
                    default:
                }
            }
        }

        // Exclude case
        // y > 4
        // h > 4
        // s > 4
        // am > 1 || pm > 1
        if (byteScore[24] > 4 || byteScore[7] > 4 || byteScore[18] > 4 || byteScore[0] > 1 || byteScore[15] > 1) {
            return false;
        }

        // Calculating the score
        // Plus 5 points for each keywords
        score += byteScore[0]  * 5; // am
        score += byteScore[3]  * 5; // d
        score += byteScore[7]  * 5; // h
        score += byteScore[12] * 5; // m
        score += byteScore[15] * 5; // pm
        score += byteScore[18] * 5; // s
        score += byteScore[24] * 5; // y

        // Addition calculation if consecutive keywords appear
        // y + m + d
        if (byteScore[24] > 0 && byteScore[12] > 0 && byteScore[3] > 0
            && byteScore[24] + byteScore[12] + byteScore[3] >= 4) {
            score += 70;
            // y + m
        } else if (byteScore[24] > 0 && byteScore[12] > 0 && byteScore[24] + byteScore[12] >= 3) {
            score += 60;
            // m + d
        } else if (byteScore[12] > 0 && byteScore[3] > 0) {
            score += 60;
            // Code is yyyy or yy
        } else if (byteScore[24] == chars.length) {
            score += 60;
        }

        // h + m + s
        if (byteScore[7] > 0 && byteScore[12] > 0 && byteScore[18] > 0
            && byteScore[7] + byteScore[12] + byteScore[18] > 3) {
            score += 70;
            // h + m
        } else if (byteScore[7] > 0 && byteScore[12] > 0) {
            score += 60;
            // m + s
        } else if (byteScore[12] > 0 && byteScore[18] > 0) {
            score += 60;
        }

        // am + pm
        if (byteScore[0] + byteScore[15] == 2) {
            score += 50;
        }

        return score >= 70;
    }

    // Find the number format in array
    private NumFmt findFmtById(int id) {
        if (numFmts == null || numFmts.isEmpty())
            return null;
        NumFmt fmt = numFmts.get(0);
        // Found it
        if (fmt.getId() == id)
            return fmt;
        if (fmt.getId() < id) {
            int size = numFmts.size();
            if (fmt.getId() < 164) {
                int index = 1;
                for (; index < size; ) {
                    if ((fmt = numFmts.get(index++)).getId() >= 164)
                        break;
                }
            }
            // The format is sorted and incremented
            // So get it by index, the index equals id subtract the first id
            int index = id - fmt.getId();
            if (index >= size) {
                fmt = null;
            } else {
                fmt = numFmts.get(id - fmt.getId());
                if (fmt.getId() == id) return fmt;
                else if (fmt.getId() < id) {
                    // Loop check forward
                    for (; index++ < size; ) {
                        NumFmt nf = numFmts.get(index);
                        if (nf.getId() == id) {
                            fmt = nf;
                            break;
                        }
                    }
                } else {
                    // Loop check backward
                    for (; index-- > 0; ) {
                        NumFmt nf = numFmts.get(index);
                        if (nf.getId() == id) {
                            fmt = nf;
                            break;
                        }
                    }
                }
            }
        } else fmt = null;
        return fmt;
    }

    /**
     * Fast test cell value is data/time value
     *
     * @param styleIndex the style index
     * @return true if the style content data format
     */
    public boolean fastTestDateFmt(int styleIndex) {
        return dateFmtCache != null && dateFmtCache.contains(styleIndex);
    }

}

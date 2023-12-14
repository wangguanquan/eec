/*
 * Copyright (c) 2017, guanquan.wang@yandex.com All Rights Reserved.
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

package org.ttzero.excel.entity.style;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.ttzero.excel.manager.TopNS;
import org.ttzero.excel.entity.I18N;
import org.ttzero.excel.entity.Storable;
import org.ttzero.excel.manager.Const;
import org.ttzero.excel.util.FileUtil;
import org.dom4j.Document;
import org.dom4j.DocumentException;
import org.dom4j.DocumentFactory;
import org.dom4j.Element;
import org.dom4j.io.SAXReader;
import org.ttzero.excel.util.StringUtil;

import java.awt.Color;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Set;
import java.util.concurrent.atomic.AtomicInteger;

import static org.ttzero.excel.util.StringUtil.isEmpty;
import static org.ttzero.excel.util.StringUtil.isNotEmpty;

/**
 * Excel样式包含格式化{@link NumFmt}、字体{@link Font}、填充{@link Fill}、边框{@link Border}、
 * 垂直对齐{@link Verticals}和水平对齐{@link Horizontals}以及自动折行组成，样式存放在共享区域供多个工作表共用，
 * 每个样式值都是由这7种值进行组合而来。全局样式库统管样式读写，添加样式时会查询当前库是否已存在，如果存在则直接返回该样式在库中的索引值，
 * 如果不存在则鼗样式添加到样式库末尾
 *
 * <p>本工具简化了样式设计，基础的4种样式由数组保存，而组合而来的单元格样式值由一个{@code int}值保存，
 * 它可以极大压缩内存消耗和快速查找，但短板是可用的样式减少，通过最多只能包含256个格式化，64个字体，填充和边框，
 * 对于日常的导出来说应该是够用的，复杂场景就需要考虑将{@code int}扩大到{@code long}</p>
 *
 * <p>由于使用2进制位点保存各个样式，所以要修改样式时必须先清除原位点上的值然后再添加新样式，否则位点在进行“或”运行时将保留全部{@code 1}
 * 致使样式错乱，可以直接使用位运算来处理样式，也可以通过{@code Styles.modify}开头的方法来修改</p>
 *
 * <blockquote><pre>
 *  Bit  | Contents
 * ------+---------
 *  0, 1 | 自动折行
 *  1, 3 | 水平对齐
 *  4, 2 | 垂直对齐
 *  6, 6 | 边框
 * 12, 6 | 填充
 * 18, 6 | 字体
 * 24. 8 | 格式化</pre></blockquote>
 *
 * @author guanquan.wang on 2017/10/13.
 */
@TopNS(prefix = "", uri = Const.SCHEMA_MAIN, value = "styleSheet")
public class Styles implements Storable {
    /**
     * LOGGER
     */
    private static final Logger LOGGER = LoggerFactory.getLogger(Styles.class);
    private final Map<Integer, Integer> map;
    private final AtomicInteger counter;
    private int[] styleIndex;
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
        counter = new AtomicInteger();
        styleIndex = new int[10];
    }

    /**
     * 向全局样式库中添加样式，添加前会先查询是否已有相同样式，如果有则直接返回该样式在全局样式库中的索引值，
     * 不存在则将该样式添加到样式库末尾
     *
     * @param s 样式值
     * @return 样式索引
     */
    public int of(int s) {
        int n = map.getOrDefault(s, -1);
        if (n == -1) {
            n = counter.getAndIncrement();
            map.put(s, n);
            if (n >= styleIndex.length) {
                styleIndex = Arrays.copyOf(styleIndex, styleIndex.length << 1);
            }
            styleIndex[n] = s;
        }
        return n;
    }

    /**
     * 通过样式索引查询样式值
     *
     * @param styleIndex 样式索引
     * @return 样式值，查找失败时返回{@code -1}
     */
    public int getStyleByIndex(int styleIndex) {
        if (styleIndex >= counter.get()) {
            return -1;
        }
        return styleIndex >= 0 ? this.styleIndex[styleIndex] : -1;
    }

    /**
     * Returns the number of styles
     *
     * @return the total styles
     */
    public int size() {
        return map.size();
    }

    public static final int INDEX_NUMBER_FORMAT = 24;
    public static final int INDEX_FONT = 18;
    public static final int INDEX_FILL = 12;
    public static final int INDEX_BORDER = 6;
    public static final int INDEX_VERTICAL = 4;
    public static final int INDEX_HORIZONTAL = 1;
    public static final int INDEX_WRAP_TEXT = 0;

    /**
     * Create a general style
     *
     * @param i18N the {@link I18N}
     * @return Styles
     */
    public static Styles create(I18N i18N) {
        Styles self = new Styles();

        self.document = createDocument();

        self.numFmts = new ArrayList<>();

        self.fonts = new ArrayList<>();
        Font font1 = new Font(i18N.get("en-font-family"), 11, Color.black);  // en
        self.addFont(font1);

        String lang = Locale.getDefault().toLanguageTag();
        // Add chinese font
        Font font2 = new Font(i18N.get("local-font-family"), 11); // cn
        if ("zh-CN".equals(lang)) {
            font2.setCharset(Charset.GB2312);
        } else if ("zh-TW".equals(lang)) {
            font2.setCharset(Charset.CHINESEBIG5);
        }
        // Other charset
        self.addFont(font2);

        self.fills = new ArrayList<>();
        self.addFill(new Fill(PatternType.none));
        self.addFill(new Fill(PatternType.gray125));

        self.borders = new ArrayList<>();
        self.addBorder(Border.parse("none"));
        self.addBorder(new Border(BorderStyle.THIN, new Color(191, 191, 191)));

        // cellXfs
        self.of(0); // General

        return self;
    }

    /**
     * 修改“边框”样式
     *
     * @param style 原样式值
     * @param newBorder 新“边框”样式
     * @return 新样式值
     */
    public int modifyBorder(int style, Border newBorder) {
        return Styles.clearBorder(style) | addBorder(newBorder);
    }

    /**
     * 修改“格式化”样式
     *
     * @param style 原样式值
     * @param newNumFmt 新“格式化”样式
     * @return 新样式值
     */
    public int modifyNumFmt(int style, NumFmt newNumFmt) {
        return Styles.clearNumFmt(style) | addNumFmt(newNumFmt);
    }

    /**
     * 修改"自动折行"样式
     *
     * @param style 原样式值
     * @param newWrapText true: 自动折行
     * @return 新样式值
     */
    public int modifyWrapText(int style, boolean newWrapText) {
        return newWrapText ? style | 1 : Styles.clearWrapText(style);
    }

    /**
     * 修改“垂直对齐”样式
     *
     * @param style 原样式值
     * @param newVertical 新“垂直对齐”样式，必须使用{@link Verticals}定义的静态值
     * @return 新样式值
     */
    public int modifyVertical(int style, int newVertical) {
        return Styles.clearVertical(style) | newVertical;
    }

    /**
     * 修改“水平对齐”样式
     *
     * @param style 原样式值
     * @param newHorizontal 新“水平对齐”样式，必须使用{@link Horizontals}定义的静态值
     * @return 新样式值
     */
    public int modifyHorizontal(int style, int newHorizontal) {
        return Styles.clearHorizontal(style) | newHorizontal;
    }

    /**
     * 修改“填充”样式
     *
     * @param style 原样式值
     * @param newFill 新“填充”样式
     * @return 新样式值
     */
    public int modifyFill(int style, Fill newFill) {
        return Styles.clearFill(style) | addFill(newFill);
    }

    /**
     * 修改“字体”样式
     *
     * @param style 原样式值
     * @param newFont 新“字体”样式
     * @return 新样式值
     */
    public int modifyFont(int style, Font newFont) {
        return Styles.clearFont(style) | addFont(newFont);
    }

    /**
     * Create a general style
     *
     * @return Styles
     */
    public static Styles forReader() {
        Styles styles = new Styles();
        styles.numFmts = new ArrayList<>();
        styles.fonts = new ArrayList<>();
        styles.fills = new ArrayList<>();
        styles.borders = new ArrayList<>();
        return styles;
    }

    /**
     * Load the style file from disk
     *
     * @param is the style {@code InputStream}
     * @return the {@link Styles} Object
     */
    public static Styles load(InputStream is) {
        // load styles.xml
        SAXReader reader = SAXReader.createDefault();
        Document document;
        try {
            document = reader.read(is);
        } catch (DocumentException e) {
            LOGGER.warn("Read the style failed and ignore the style to continue.", e);
            Styles self = forReader();
            // Add a default font
            self.addFont(new Font("Arial", 11, Color.black));
            return self;
        }

        Styles self = new Styles();
        Element root = document.getRootElement();

        // Parse Number format
        self.numFmts = NumFmt.domToNumFmt(root);

        // Parse Fonts
        self.fonts = Font.domToFont(root);

        // Parse Fills
        self.fills = Fill.domToFill(root);

        // Parse Borders
        self.borders = Border.domToBorder(root);

        // Cell xf
        Element cellXfs = root.element("cellXfs");
        List<Element> sub = cellXfs.elements();
        int i = 0;
        for (Element e : sub) {
            int style = 0;
            // NumFmt
            String numFmtId = getAttr(e, "numFmtId"); // applyNumberFormat = getAttr(e, "applyNumberFormat");
            if (StringUtil.isNotEmpty(numFmtId) && !"0".equals(numFmtId)) {
                style |= Integer.parseInt(numFmtId) << INDEX_NUMBER_FORMAT;
            }
            // Font
            String fontId = getAttr(e, "fontId"); // applyFont = getAttr(e, "applyFont");
            if (StringUtil.isNotEmpty(fontId) && !"0".equals(fontId)) {
                style |= Integer.parseInt(fontId) << INDEX_FONT;
            }
            // Fill
            String fillId = getAttr(e, "fillId"); // applyFill = getAttr(e, "applyFill");
            if (StringUtil.isNotEmpty(fillId) && !"0".equals(fillId)) {
                style |= Integer.parseInt(fillId) << INDEX_FILL;
            }
            // Border
            String borderId = getAttr(e, "borderId"); // applyBorder = getAttr(e, "applyBorder");
            if (StringUtil.isNotEmpty(borderId) && !"0".equals(borderId)) {
                style |= Integer.parseInt(borderId) << INDEX_BORDER;
            }
            // Alignment
            Element alignment = e.element("alignment");
            if (alignment != null) {
                String horizontal = getAttr(alignment, "horizontal");
                int index;
                if (StringUtil.isNotEmpty(horizontal) && (index = StringUtil.indexOf(Horizontals._names, horizontal)) >= 0) {
                    style |= index << INDEX_HORIZONTAL;
                }
                String vertical = getAttr(alignment, "vertical");
                if (StringUtil.isNotEmpty(vertical) && (index = StringUtil.indexOf(Verticals._names, vertical)) >= 0) {
                    style |= index << INDEX_VERTICAL;
                } else style |= Verticals.BOTTOM;
                String wrapText = getAttr(alignment, "wrapText");
                style |= ("1".equals(wrapText) || "true".equalsIgnoreCase(wrapText) ? 1 : 0) << INDEX_WRAP_TEXT;
            }
            self.map.put(style, i);
            if (i >= self.styleIndex.length) {
                self.styleIndex = Arrays.copyOf(self.styleIndex, self.styleIndex.length << 1);
            }
            self.styleIndex[i] = style;
            i++;
        }
        self.counter.set(i);
        // Test number format
        for (Integer styleIndex : self.map.values()) {
            self.isDate(styleIndex);
        }

        return self;
    }

    /**
     * 添加“格式化”，对格式化串去重处理
     *
     * <p>返回样式值中“格式化”部分的2进制值，拿到这个值后可以与其它部分值进行“或”运算以组成最终的样式值</p>
     *
     * @param numFmt 格式化{@link NumFmt}
     * @return 样式值中“格式化”部分的2进制值
     */
    public final int addNumFmt(NumFmt numFmt) {
        // All indexes from 0 to 175 are reserved for built-in formats.
        // The first user-defined format starts at 176.
        if (numFmt.getId() < 0 || numFmt.getId() >= 176) {
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
                } else {
                    numFmt.setId(numFmts.get(i).getId());
                }
            }
        }
        return numFmt.getId() << INDEX_NUMBER_FORMAT;
    }

    /**
     * 添加“字体”
     *
     * <p>返回样式值中“字体”部分的2进制值，拿到这个值后可以与其它部分值进行“或”运算以组成最终的样式值</p>
     *
     * @param font 字体{@link Font}
     * @return 样式值中“字体”部分的2进制值
     */
    public final int addFont(Font font) {
        if (isEmpty(font.getName())) {
            throw new IllegalArgumentException("Font name not support.");
        }
        int i = fonts.indexOf(font);
        if (i <= -1) {
            i = fonts.size();
            fonts.add(font);
        }
        return i << INDEX_FONT;
    }

    /**
     * 添加“填充”
     *
     * <p>返回样式值中“填充”部分的2进制值，拿到这个值后可以与其它部分值进行“或”运算以组成最终的样式值</p>
     *
     * @param fill 填充{@link Font}
     * @return 样式值中“填充”部分的2进制值
     */
    public final int addFill(Fill fill) {
        int i = fills.indexOf(fill);
        if (i <= -1) {
            i = fills.size();
            fills.add(fill);
        }
        return i << INDEX_FILL;
    }

    /**
     * 添加“边框”
     *
     * <p>返回样式值中“边框”部分的2进制值，拿到这个值后可以与其它部分值进行“或”运算以组成最终的样式值</p>
     *
     * @param border 边框{@link Border}
     * @return 样式值中“边框”部分的2进制值
     */
    public final int addBorder(Border border) {
        int i = borders.indexOf(border);
        if (i <= -1) {
            i = borders.size();
            borders.add(border);
        }
        return i << INDEX_BORDER;
    }

    public static int[] unpack(int style) {
        int[] styles = new int[7];
        styles[0] = style >>> INDEX_NUMBER_FORMAT;
        styles[1] = style << 8 >>> (INDEX_FONT + 8);
        styles[2] = style << 14 >>> (INDEX_FILL + 14);
        styles[3] = style << 20 >>> (INDEX_BORDER + 20);
        styles[4] = style << 26 >>> (INDEX_VERTICAL + 26);
        styles[5] = style << 28 >>> (INDEX_HORIZONTAL + 28);
        styles[6] = style << 31 >>> (INDEX_WRAP_TEXT + 31);
        return styles;
    }

    public static int pack(int[] styles) {
        return styles[0] << INDEX_NUMBER_FORMAT
            | styles[1] << INDEX_FONT
            | styles[2] << INDEX_FILL
            | styles[3] << INDEX_BORDER
            | styles[4] << INDEX_VERTICAL
            | styles[5] << INDEX_HORIZONTAL
            | styles[6] << INDEX_WRAP_TEXT
            ;
    }

    private static final String[] attrNames = {
            "numFmtId"
            , "fontId"
            , "fillId"
            , "borderId"
            , "vertical"
            , "horizontal"
            , "wrapText"
            , "applyNumberFormat"
            , "applyFont"
            , "applyFill"
            , "applyBorder"
            , "applyAlignment"
    };

    /**
     * Write style to disk
     *
     * @param styleFile the storage path
     * @throws IOException if I/O error occur
     */
    @Override
    public void writeTo(Path styleFile) throws IOException {
        if (document == null) document = createDocument();
        Element root = document.getRootElement();

        // Number format
        if (!numFmts.isEmpty()) {
            Element element = document.getRootElement().element("numFmts");
            element.attribute("count").setValue(String.valueOf(numFmts.size()));
            for (NumFmt numFmt : numFmts) numFmt.toDom(element);
        }

        // Font
        if (!fonts.isEmpty()) {
            Element element = document.getRootElement().element("fonts");
            element.attribute("count").setValue(String.valueOf(fonts.size()));
            for (Font font : fonts) font.toDom(element);
        }

        // Fill
        if (!fills.isEmpty()) {
            Element element = document.getRootElement().element("fills");
            element.attribute("count").setValue(String.valueOf(fills.size()));
            for (Fill fill : fills) fill.toDom(element);
        }

        // Border
        if (!borders.isEmpty()) {
            Element element = document.getRootElement().element("borders");
            element.attribute("count").setValue(String.valueOf(borders.size()));
            for (Border border : borders) border.toDom(element);
        }

        Element cellXfs = root.element("cellXfs").addAttribute("count", String.valueOf(map.size()));

        for (int i = 0, len = counter.get(); i < len; i++) {
            int[] styles = unpack(styleIndex[i]);

            Element newXf = cellXfs.addElement("xf");
            newXf.addAttribute(attrNames[0], String.valueOf(styles[0]))
                .addAttribute(attrNames[1], String.valueOf(styles[1]))
                .addAttribute(attrNames[2], String.valueOf(styles[2]))
                .addAttribute(attrNames[3], String.valueOf(styles[3]))
                .addAttribute("xfId", "0")
            ;
            int start = 7;
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
            if ((styles[4] | styles[5] | styles[6]) > 0) {
                newXf.addAttribute(attrNames[start + 4], "1");
            }

            Element subEle = newXf.addElement("alignment").addAttribute(attrNames[4], Verticals._names[styles[4]]);
            if (styles[5] > 0) {
                subEle.addAttribute(attrNames[5], Horizontals._names[styles[5]]);
            }
            if (styles[6] > 0) {
                subEle.addAttribute(attrNames[6], "1");
            }
        }

        FileUtil.writeToDiskNoFormat(document, styleFile);
    }

    public static Document createDocument() {
        DocumentFactory factory = DocumentFactory.getInstance();
        TopNS ns = Styles.class.getAnnotation(TopNS.class);
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
            .addAttribute("name", "常规")
            .addAttribute("xfId", "0");

        return factory.createDocument(rootElement);
    }

    ////////////////////////clear style///////////////////////////////

    /**
     * 清除样式中的“格式化”
     *
     * @param style 样式值
     * @return 清除“格式化”后的样式
     */
    public static int clearNumFmt(int style) {
        return style & (-1 >>> 32 - INDEX_NUMBER_FORMAT);
    }

    /**
     * 清除样式中的“字体”
     *
     * @param style 样式值
     * @return 清除“字体”后的样式
     */
    public static int clearFont(int style) {
        return style & ~((-1 >>> 32 - (INDEX_NUMBER_FORMAT - INDEX_FONT)) << INDEX_FONT);
    }

    /**
     * 清除样式中的“填充”
     *
     * @param style 样式值
     * @return 清除“填充”后的样式
     */
    public static int clearFill(int style) {
        return style & ~((-1 >>> 32 - (INDEX_FONT - INDEX_FILL)) << INDEX_FILL);
    }

    /**
     * 清除样式中的“边框”
     *
     * @param style 样式值
     * @return 清除“边框”后的样式
     */
    public static int clearBorder(int style) {
        return style & ~((-1 >>> 32 - (INDEX_FILL - INDEX_BORDER)) << INDEX_BORDER);
    }

    /**
     * 清除样式中的“垂直对齐”
     *
     * @param style 样式值
     * @return 清除“垂直对齐”后的样式
     */
    public static int clearVertical(int style) {
        return style & ~((-1 >>> 32 - (INDEX_BORDER - INDEX_VERTICAL)) << INDEX_VERTICAL);
    }

    /**
     * 清除样式中的“水平对齐”
     *
     * @param style 样式值
     * @return 清除“水平对齐”后的样式
     */
    public static int clearHorizontal(int style) {
        return style & ~((-1 >>> 32 - (INDEX_VERTICAL - INDEX_HORIZONTAL)) << INDEX_HORIZONTAL);
    }

    /**
     * 清除样式中的“自动折行”
     *
     * @param style 样式值
     * @return 清除“自动折行”后的样式
     */
    public static int clearWrapText(int style) {
        return style & ~(-1 >>> 32 - (INDEX_HORIZONTAL - INDEX_WRAP_TEXT));
    }

    ////////////////////////reset style/////////////////////////////
    @Deprecated
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
        return (1 << INDEX_BORDER) | (1 << INDEX_FONT) | Horizontals.CENTER;
    }

    public static int defaultStringBorderStyle() {
        return  (1 << INDEX_BORDER) | (1 << INDEX_FONT) | Horizontals.LEFT;
    }

    public static int defaultIntBorderStyle() {
        return (1 << INDEX_BORDER) | (1 << INDEX_FONT) | Horizontals.RIGHT;
    }

    public static int defaultDoubleBorderStyle() {
        return (1 << INDEX_BORDER) | (1 << INDEX_FONT) | Horizontals.RIGHT;
    }

    ////////////////////////default style/////////////////////////////
    public static int defaultCharStyle() {
        return (1 << INDEX_FONT) | Horizontals.CENTER;
    }

    public static int defaultStringStyle() {
        return (1 << INDEX_FONT) | Horizontals.LEFT;
    }

    public static int defaultIntStyle() {
        return (1 << INDEX_FONT) | Horizontals.RIGHT;
    }

    public static int defaultDoubleStyle() {
        return (1 << INDEX_FONT) | Horizontals.RIGHT;
    }

    ////////////////////////////////Check style////////////////////////////////

    /**
     * 判断样式是否包含“格式化"
     *
     * @param style 样式值
     * @return true: 包含”格式化“
     */
    public static boolean hasNumFmt(int style) {
        return style >>> INDEX_NUMBER_FORMAT != 0;
    }

    /**
     * 判断样式是否包含“字体"
     *
     * @param style 样式值
     * @return true: 包含”字体“
     */
    public static boolean hasFont(int style) {
        return true; // Font is required
    }

    /**
     * 判断样式是否包含“填充"
     *
     * @param style 样式值
     * @return true: 包含”填充“
     */
    public static boolean hasFill(int style) {
        return style << 14 >>> (INDEX_FILL + 14) != 0;
    }

    /**
     * 判断样式是否包含“边框"
     *
     * @param style 样式值
     * @return true: 包含”边框“
     */
    public static boolean hasBorder(int style) {
        return style << 20 >>> (INDEX_BORDER + 20) != 0;
    }

    /**
     * 判断样式是否包含“垂直对齐"
     *
     * @param style 样式值
     * @return true: 包含”垂直对齐“（非默认）
     */
    public static boolean hasVertical(int style) {
        return style << 26 >>> (INDEX_VERTICAL + 26) != 0;
    }

    /**
     * 判断样式是否包含“水平对齐"
     *
     * @param style 样式值
     * @return true: 包含”水平对齐“（非默认）
     */
    public static boolean hasHorizontal(int style) {
        return style << 28 >>> (INDEX_HORIZONTAL + 28) != 0;
    }

    /**
     * 判断样式是否自动折行
     *
     * @param style 样式值
     * @return true: 自动折行
     */
    public static boolean hasWrapText(int style) {
        return (style & 1) > 0;
    }

    ////////////////////////////////To object//////////////////////////////////

    /**
     * 获取样式中的格式化值
     *
     * @param style 样式值
     * @return 格式化或 {@code null}
     */
    public NumFmt getNumFmt(int style) {
        int n = style >>> INDEX_NUMBER_FORMAT;
        if (n <= 0) return null;
        if (n < 176) return BuiltInNumFmt.get(n);
        for (NumFmt e : numFmts) {
            if (e.id == n) return e;
        }
        return null;
    }

    /**
     * 获取样式中的格式化
     *
     * @param style 样式值
     * @return 当前样式包含的格式化，不含格式化时返回{@code null}
     */
    public Fill getFill(int style) {
        return fills.get(style << 14 >>> (INDEX_FILL + 14));
    }

    /**
     * 获取样式中的字体
     *
     * @param style 样式值
     * @return 当前样式包含的字体，样式一定包含字体
     */
    public Font getFont(int style) {
        return fonts.get(Math.max(0, style << 8 >>> (INDEX_FONT + 8)));
    }

    /**
     * 获取样式中的边框
     *
     * @param style 样式值
     * @return 当前样式包含的边框，不含边框时返回{@code null}
     */
    public Border getBorder(int style) {
        return borders.get(style << 20 >>> (INDEX_BORDER + 20));
    }

    /**
     * 获取样式中的垂直对齐，参考范围{@link Verticals}
     *
     * @param style 样式值
     * @return 当前样式包含的垂直对齐
     */
    public int getVertical(int style) {
        return style << 26 >>> (INDEX_VERTICAL + 26) << INDEX_VERTICAL;
    }

    /**
     * 获取样式中的水平对齐，参考范围{@link Horizontals}
     *
     * @param style 样式值
     * @return 当前样式包含的水平对齐
     */
    public int getHorizontal(int style) {
        return style << 28 >>> (INDEX_HORIZONTAL + 28) << INDEX_HORIZONTAL;
    }

    /**
     * 获取样式中自动折行标记，标记为{@code 1}时表示自动折行
     *
     * @param style 样式值
     * @return 1: 自动折行
     */
    public int getWrapText(int style) {
        return style & 1;
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
     * Parse color tag
     *
     * @param element color tag
     * @return awt.Color or null
     */
    public static Color parseColor(Element element) {
        if (element == null) return null;
        String rgb = getAttr(element, "rgb"), indexed = getAttr(element, "indexed")
            , auto = getAttr(element, "auto"), theme = getAttr(element, "theme");
        Color c = null;
        // Standard Alpha Red Green Blue color value (ARGB).
        if (StringUtil.isNotEmpty(rgb)) {
            c = ColorIndex.toColor(rgb);
        }
        // Indexed color value. Only used for backwards compatibility.
        // References a color in indexedColors.
        else if (StringUtil.isNotEmpty(indexed)) {
            // if indexed greater than 64 means auto.
            c = new BuildInColor(Integer.parseInt(indexed));
        }
        // A boolean value indicating the color is automatic and system color dependent.
        else if ("1".equals(auto) || "true".equalsIgnoreCase(auto)) {
            c = new BuildInColor(64);
        }
        // Theme colors
        else if (StringUtil.isNotEmpty(theme)) {
            int t = 0;
            try {
                t = Integer.parseInt(theme);
            } catch (NumberFormatException ex) { }
            if (t < 0 || t > 11) {
                LOGGER.warn("Unknown theme color index {}", t);
                t = 0;
            }
            Color themeColor = ColorIndex.themeColors[t];
            String tint = getAttr(element, "tint");
            c = HlsColor.calculateColor(themeColor, tint);
        }
        return c;
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

        if (styleIndex > counter.get()) return false;
        int style = this.styleIndex[styleIndex];
        int nf = style >> INDEX_NUMBER_FORMAT & 0xFF;

        // No number format
        if (nf == 0) return false;

        boolean isDate;
        NumFmt numFmt;

        // Test by numFmt code
        if (!(isDate = isBuildInDateFormat(nf))
            && (numFmt = findFmtById(nf)) != null
            && (isNotEmpty(numFmt.getCode()))) {
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

    // All indexes from 0 to 163 are reserved for built-in formats.
    // The first user-defined format starts at 164.
    private static boolean isBuildInDateFormat(int nf) {
        return nf < 164 && (nf >= 14 && nf <= 22
            || nf >= 27 && nf <= 36
            || nf >= 45 && nf <= 47
            || nf >= 50 && nf <= 58
            || nf == 81);
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
        NumFmt fmt = numFmts.get(numFmts.size() - 1);
        if (fmt.getId() > id) {
            int n = Collections.binarySearch(numFmts, new NumFmt(id, null));
            return n >= 0 ? numFmts.get(n) : null;
        }
        return fmt.getId() == id ? fmt : null;
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

    /**
     * Append a date format back into cache
     *
     * @param xf the XFRecord id
     */
    public void addDateFmtCache(int xf) {
        if (dateFmtCache == null) dateFmtCache = new HashSet<>();
        dateFmtCache.add(xf);
    }

    /**
     * Converts a <code>String</code> to an integer and returns the
     * specified opaque <code>Color</code>. This method handles string
     * formats that are used to represent octal and hexadecimal numbers.
     *
     * @param v hexadecimal numbers or color name
     * @return the new <code>Color</code> object.
     * @throws IllegalArgumentException if convert failed.
     */
    public static Color toColor(String v) {
        Color color;
        if (v.charAt(0) == '#') {
            try {
                color = Color.decode(v);
            } catch (NumberFormatException e) {
                throw new IllegalArgumentException("Color \"" + v + "\" not support.");
            }
        } else {
            try {
                Field field = Color.class.getDeclaredField(v);
                color = (Color) field.get(null);
            } catch (NoSuchFieldException | IllegalAccessException e) {
                throw new IllegalArgumentException("Color \"" + v + "\" not support.");
            }
        }
        return color;
    }
}

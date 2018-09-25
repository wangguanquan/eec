package net.cua.export.entity.e7.style;

import net.cua.export.annotation.TopNS;
import net.cua.export.manager.Const;
import net.cua.export.tmap.TIntIntHashMap;
import net.cua.export.util.FileUtil;
import net.cua.export.util.StringUtil;
import org.dom4j.Document;
import org.dom4j.DocumentFactory;
import org.dom4j.Element;

import java.awt.Color;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;

/**
 * 每个style由一个int值组成
 * 0~ 8位 numFmt
 * 8~14位 font
 * 14~20位 fill
 * 20~26位 border
 * 26~29位 vertical
 * 29~32位 horizontal
 * Created by guanquan.wang on 2017/10/13.
 */
@TopNS(prefix = "", uri = Const.SCHEMA_MAIN, value = "styleSheet")
public class Styles {

    private TIntIntHashMap map;
    private Document document;

    List<Font> fonts;
    List<NumFmt> numFmts;
    List<Fill> fills;
    List<Border> borders;

    private Styles() {
        map = new TIntIntHashMap();
    }

    /**
     * 根据位编码找到style下标
     *
     * @param s 位编码
     * @return
     */
    public int of(int s) {
        int n = map.get(s);
        if (n == 0) {
            n = addStyle(s);
            map.put(s, n);
        }
        return n;
    }

    static final int INDEX_NUMBER_FORMAT = 24;
    static final int INDEX_FONT = 18;
    static final int INDEX_FILL = 12;
    static final int INDEX_BORDER = 6;
    static final int INDEX_VERTICAL = 3;
    static final int INDEX_HORIZONTAL = 0;

    /**
     * create general style
     *
     * @return
     */
    public static final Styles create() {
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
        // cellStyles
        Element cellStyles = rootElement.addElement("cellStyles").addAttribute("count", "1");
        cellStyles.addElement("cellStyle")
                .addAttribute("builtinId", "0")
                .addAttribute("name", "常规") // TODO I18N
                .addAttribute("xfId", "0");

        self.document = factory.createDocument(rootElement);

        self.numFmts = new ArrayList<>();
        self.addNumFmt(new NumFmt("yyyy\\-mm\\-dd"));
        self.addNumFmt(new NumFmt("yyyy\\-mm\\-dd\\ hh:mm:ss"));

        self.fonts = new ArrayList<>();
        Font font1 = new Font("宋体", 9, Color.black);  // en
        font1.setFamily(2);
        font1.setScheme("minor");
        self.addFont(font1);

        Font font2 = new Font("宋体", 11); // cn
        font2.setFamily(3);
        font2.setScheme("minor");
        font2.setCharset(Charset.GB2312);
        self.addFont(font2);

        self.fills = new ArrayList<>();
        self.addFill(new Fill(PatternType.none));
        self.addFill(new Fill(PatternType.gray125));
        self.addFill(Fill.parse("#fcfaf2"));

        self.borders = new ArrayList<>();
        self.addBorder(Border.parse("none"));
        self.addBorder(Border.parse("thin black"));

        // cellXfs
        self.of(0); // General

        return self;
    }

    /**
     * add number format
     *
     * @param numFmt
     * @return
     */
    public synchronized final int addNumFmt(NumFmt numFmt) {
        // check and search default code
        if (numFmt.getId() < 0) {
            if (StringUtil.isEmpty(numFmt.getCode())) {
                throw new NullPointerException("NumFmt code");
            }
            int index = DefaultNumFmt.indexOf(numFmt.getCode());
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
     * add font
     *
     * @param font
     * @return
     */
    public synchronized final int addFont(Font font) {
        if (StringUtil.isEmpty(font.getName())) {
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
     * add fill
     *
     * @param fill
     * @return
     */
    public synchronized final int addFill(Fill fill) {
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
     * add border
     *
     * @param border
     * @return
     */
    public synchronized final int addBorder(Border border) {
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

    public static int[] unpack(int style) {
        int[] styles = new int[6];
        styles[0] = style >>> INDEX_NUMBER_FORMAT;
        styles[1] = style << 8 >>> (INDEX_FONT + 8);
        styles[2] = style << 14 >>> (INDEX_FILL + 14);
        styles[3] = style << 20 >>> (INDEX_BORDER + 20);
        styles[4] = style << 26 >>> (INDEX_VERTICAL + 26);
        styles[5] = style << 29 >>> (INDEX_HORIZONTAL + 29);
        return styles;
    }

    public static int pack(int[] styles) {
        return styles[0] << INDEX_NUMBER_FORMAT
                | styles[1] << INDEX_FONT
                | styles[2] << INDEX_FILL
                | styles[3] << INDEX_BORDER
                | styles[4] << INDEX_VERTICAL
                | styles[5] << INDEX_HORIZONTAL
                ;
    }

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

        String[] attrNames = {"numFmtId", "fontId", "fillId", "borderId", "vertical", "horizontal"
                , "applyNumberFormat", "applyFont", "applyFill", "applyBorder", "applyAlignment"};
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

    public void writeTo(Path styleFile) throws IOException {
        if (document != null) { // Not null
            FileUtil.writeToDisk(document, styleFile);
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
        return style <<26>>>(INDEX_VERTICAL +26);
    }

}

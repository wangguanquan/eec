package net.cua.export.entity.e7;

import net.cua.export.annotation.TopNS;
import net.cua.export.manager.Const;
import net.cua.export.tmap.TIntIntHashMap;
import net.cua.export.util.FileUtil;
import net.cua.export.util.StringUtil;
import org.dom4j.Attribute;
import org.dom4j.Document;
import org.dom4j.DocumentException;
import org.dom4j.Element;
import org.dom4j.io.SAXReader;

import java.io.File;
import java.io.InputStream;
import java.util.Arrays;
import java.util.Iterator;
import java.util.List;

/**
 * 每个style由一个int值组成
 * 0~ 8位 numFmt
 * 8~14位 font
 * 14~20位 fill
 * 20~26位 border
 * 26~29位 vertical
 * 29~32位 horizontal
 * Created by wanggq on 2017/10/13.
 */
@TopNS(prefix = "", uri = Const.SCHEMA_MAIN, value = "styleSheet")
public class Styles {

    private TIntIntHashMap map;
    private Document document;

//    private static class Holder { // 每个workbook一个实例
//        private static final Styles INSTANCE = new Styles();
//    }
//
//    public static final Styles getInstance() {
//        return Holder.INSTANCE;
//    }

    Styles() {
        map = new TIntIntHashMap();
    }

    /**
     * 根据位编码找到style下标
     *
     * @param s 位编码
     * @return
     */
    public int of(int s) {
        if (s == 0) return s;
        int n = map.get(s); // not found return 0(default style)
        if (n == 0) {
            n = addStyle(s);
            map.put(s, n);
        }
        return n;
    }

    String[] attrNames = {"numFmtId", "fontId", "fillId", "borderId", "vertical", "horizontal"};
    static final int[] move_left = {24, 18, 12, 6, 3, 0};

    public Styles load(InputStream is) {
        map.put(0, 0);
        map.put(1, 1);
        SAXReader reader = new SAXReader();
        try {
            document = reader.read(is);
        } catch (DocumentException e) {
            e.printStackTrace();
            // TODO read style file fail.
            return this;
        }
        Element root = document.getRootElement();
        Element cellXfs = root.element("cellXfs");
        Iterator<Element> elementIterator = cellXfs.elementIterator();
        int n = 0;
        while (elementIterator.hasNext()) {
            Element xf = elementIterator.next();
            if (++n <= 2) continue;
            Element alignment = xf.element("alignment");
            int c = 0;
            for (int i = 0; i < attrNames.length; i++) {
                Attribute attr = xf.attribute(attrNames[i]);
                if (attr == null && alignment != null) {
                    attr = alignment.attribute(attrNames[i]);
                }
                if (attr != null) {
                    String attrValue = attr.getValue();
                    int v;
                    if (i < 4) {
                        v = Integer.parseInt(attrValue);
                    } else if (i == 4) {
                        v = Verticals.valueOf(attrValue);
                    } else {
                        v = Horizontals.valueOf(attrValue);
                    }
//                    System.out.print(v + "   ");
                    c |= (v << move_left[i]);
                }
            }

//            System.out.println(c + " : " + (n - 1));
            map.put(c, n - 1);
        }
        return this;
    }

    public int[] unpackStyle(int style) {
        int[] styles = new int[6];
        styles[0] = style >>> move_left[0];
        styles[1] = style << 8 >>> move_left[1] + 8;
        styles[2] = style << 14 >>> move_left[2] + 14;
        styles[3] = style << 20 >>> move_left[3] + 20;
        styles[4] = style << 26 >>> move_left[4] + 26;
        styles[5] = style << 29 >>> move_left[5] + 29;
        return styles;
    }

    @Override
    public String toString() {
        StringBuilder buf = new StringBuilder();
        buf.append("<cellXfs count=\"").append(map.size()).append("\">\n");
        int[] keys = map.keys(), values = map.values();
        for (int i = 0; i < keys.length; i++) {
            int k = keys[indexOf(values, i)];
            int[] styles = unpackStyle(k);
//            System.out.println(styles[0] + "   " + styles[1] + "   " + styles[2] + "   " + styles[3] + "   " + styles[4] + "   " + styles[5]);
            buf.append("<xf numFmtId=\"").append(styles[0]).append("\"")
                    .append(" fontId=\"").append(styles[1]).append("\"")
                    .append(" fillId=\"").append(styles[2]).append("\"")
                    .append(" borderId=\"").append(styles[3]).append("\"")
            ;
            if (styles[0] > 0) {
                buf.append(" applyNumberFormat=\"1\"");
            }
            if (styles[1] > 0) {
                buf.append(" applyFont=\"1\"");
            }
            if (styles[2] > 0) {
                buf.append(" applyFill=\"1\"");
            }
            if (styles[3] > 0) {
                buf.append(" applyBorder=\"1\"");
            }
            if ((styles[4] | styles[5]) > 0) {
                buf.append(" applyAlignment=\"1\"");
            }

            buf.append(">\n   <alignment vertical=\"").append(Verticals.of(styles[4])).append("\"");
            if (styles[5] > 0) {
                int horizontal = styles[5];
                if (k == 1) horizontal = 3;
                buf.append(" horizontal=\"").append(Horizontals.of(horizontal)).append("\"");
            }
            buf.append(" />\n</xf>\n");
        }
        buf.append("</cellXfs>");
        return buf.toString();
    }

    private synchronized int addStyle(int s) {
        if (document == null) return 0;
        int[] styles = unpackStyle(s);
//        System.out.println(styles[0] + "   " + styles[1] + "   " + styles[2] + "   " + styles[3] + "   " + styles[4] + "   " + styles[5]);
        Element root = document.getRootElement();
        Element cellXfs = root.element("cellXfs");
        int count = Integer.valueOf(cellXfs.attributeValue("count"));
        int n = cellXfs.elements().size();
        Element newXf = cellXfs.addElement("xf");
        newXf.addAttribute(attrNames[0], String.valueOf(styles[0]))
                .addAttribute(attrNames[1], String.valueOf(styles[1]))
                .addAttribute(attrNames[2], String.valueOf(styles[2]))
                .addAttribute(attrNames[3], String.valueOf(styles[3]))
                .addAttribute("xfId", "0")
        ;
        if (styles[0] > 0) {
            newXf.addAttribute("applyNumberFormat", "1");
        }
        if (styles[1] > 0) {
            newXf.addAttribute("applyFont", "1");
        }
        if (styles[2] > 0) {
            newXf.addAttribute("applyFill", "1");
        }
        if (styles[3] > 0) {
            newXf.addAttribute("applyBorder", "1");
        }
        if ((styles[4] | styles[5]) > 0) {
            newXf.addAttribute("applyAlignment", "1");
        }

        Element subEle = newXf.addElement("alignment").addAttribute(attrNames[4], Verticals.of(styles[4]));
        if (styles[5] > 0) {
            subEle.addAttribute(attrNames[5], Horizontals.of(styles[5]));
        }
        cellXfs.addAttribute("count", String.valueOf(count+1));
        return n;
    }

    public void writeTo(File styleFile) {
        if (document != null) {
            FileUtil.writeToDisk(document, styleFile.getPath());
        } else {
            FileUtil.copyFile(getClass().getClassLoader().getResourceAsStream("template/styles.xml"), styleFile);
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

    ////////////////////////clear style//////////////////////////////
    public static int clearNumfmt(int style) {
        return style & (-1 >>> 32 - move_left[0]);
    }

    public static int clearFont(int style) {
        return style & ~((-1 >>> 32 - (move_left[0] - move_left[1])) << move_left[1]);
    }

    public static int clearFill(int style) {
        return style & ~((-1 >>> 32 - (move_left[1] - move_left[2])) << move_left[2]);
    }

    public static int clearBorder(int style) {
        return style & ~((-1 >>> 32 - (move_left[2] - move_left[3])) << move_left[3]);
    }

    public static int clearVertical(int style) {
        return style & ~((-1 >>> 32 - (move_left[3] - move_left[4])) << move_left[4]);
    }

    public static int clearHorizontal(int style) {
        return style & ~(-1 >>> 32 - (move_left[4] - move_left[5]));
    }

    ///////////////////default style///////////////////

    public static int defaultStringStyle() {
        return Styles.Fonts.BLACK_GB2312_WRYH_11| Styles.Borders.THIN_BLACK| Styles.Horizontals.LEFT;
    }
    public static int defaultIntStyle() {
        return Styles.NumFmts.PADDING_INT|Styles.Fonts.BLACK_ASCII_CONSOLAS_11| Styles.Borders.THIN_BLACK| Styles.Horizontals.RIGHT;
    }
    public static int defaultDateStyle() {
        return Styles.NumFmts.DATE|Styles.Fonts.BLACK_ASCII_CONSOLAS_11| Styles.Borders.THIN_BLACK| Styles.Horizontals.CENTER_CONTINUOUS;
    }
    public static int defaultTimestampStyle() {
        return Styles.NumFmts.DATE_TIME| Styles.Fonts.BLACK_ASCII_CONSOLAS_11| Styles.Borders.THIN_BLACK| Styles.Horizontals.CENTER_CONTINUOUS;
    }
    public static int defaultDoubleStyle() {
        return Styles.NumFmts.PADDING_DOUBLE | Styles.Fonts.BLACK_ASCII_CONSOLAS_11 | Styles.Borders.THIN_BLACK | Styles.Horizontals.RIGHT;
    }


    public static final class NumFmts {
        public static final int GENERAL = 0 // General
                , INT = 1 << move_left[0] // 0
                , DOUBLE = 2 << move_left[0] // 0.00

                , MARK_INT = 3 << move_left[0] // #,##0
                , MARK_DOUBLE = 4 << move_left[0] // #,##0.00

                , PERCENTAGE_INT = 9 << move_left[0] // 0%
                , PERCENTAGE_DOUBLE = 10 << move_left[0] // 0.00%

                , PADDING_MARK_INT = 38 << move_left[0] // #,##0_);[Red](#,##0)
                , PADDING_MARK_DOUBLE = 178 << move_left[0] // #,##0.00_);[Red](#,##0.00)

                , DOUBLE_3 = 179 << move_left[0] // 0.000
                , PADDING_DOUBLE_3 = 180 << move_left[0] // 0.000_);[Red](0.000)

                , PADDING_PERCENTAGE_INT = 188 << move_left[0] // 0%_);[Red](0%) 百分比默认样式
                , PADDING_PERCENTAGE_DOUBLE = 181 << move_left[0] // 0.00%_);[Red](0.00%) 百分比默认样式

                , PADDING_INT = 176 << move_left[0] // 0_);[Red](0) 整数默认样式
                , PADDING_DOUBLE = 177 << move_left[0] // 0.00_);[Red](0.00) 小数默认样式

                , YEN_INT =  182 << move_left[0] // ¥0
                , YEN_DOUBLE = 183 << move_left[0] // ¥0.00

                , PADDING_YEN_INT = 184 << move_left[0] // ¥0_);[Red](¥0)
                , PADDING_YEN_DOUBLE =  185 << move_left[0] // ¥0.00_);[Red](¥0.00) 货币默认样式

                , DATE = 186 << move_left[0] // yyyy-mm-dd  date默认样式
                , DATE_TIME = 187 << move_left[0] // yyyy-mm-dd hh:mm:ss timestamp默认样式
                ;
    }

    public static class Fonts {
        public static final int BLACK_ASCII_SONG_11 = 0 // black|default|宋体|11
                , BLACK_GB2312_SONG_9 = 1 << move_left[1] // black|gb2312|宋体|8
                , WHITE_GB2312_WRYH_11_B = 2 << move_left[1] // white|gb2312|微软雅黑|11|加粗 列表头默认字体
                , BLACK_ASCII_CONSOLAS_11 = 3 << move_left[1] // 正文数字默认字体
                , BLACK_GB2312_WRYH_11 = 4 << move_left[1] // 正文汉字默认
                , RED_ASCII_CONSOLAS_11 = 5 << move_left[1] // 正文数字标红字体
                , RED_ASCII_CONSOLAS_11_B = 6 << move_left[1] // 正文数字标红加粗
                , BLACK_ASCII_CONSOLAS_11_I = 7 << move_left[1] // 数字斜体
                ;
    }

    public static class Fills {
        public static final int NONE = 0 // 无填充
                , GRAY125 = 1 << move_left[2]
                , FF666699 = 2 << move_left[2] // 列表头背景色
                , RED = 3 << move_left[2] // 红色背景色
                , YELLOW = 4 << move_left[2] // 黄色北景色
                ;
    }

    public static class Borders {
        public static final int NONE = 0 // 无边框
                , THIN_BLACK = 1 << move_left[3] // 黑色连续边框
                ;
    }

    public static final class Verticals {
        public static final int CENTER = 0 // Align Center
                , BOTTOM = 1 << move_left[4] // Align Bottom
                , TOP = 2 << move_left[4]   // Align Top
                , BOTH = 3 << move_left[4] // Vertical Justification
                ;

        private static final String[] _names = {"center", "bottom", "top", "both"};
        public static int valueOf(String name) {
            return StringUtil.indexOf(_names, name);
        }

        public static String of(int n) {
            return _names[n];
        }
    }

    public static final class Horizontals {
        public static final int GENERAL = 0 // General Horizontal Alignment( Text data is left-aligned. Numbers, dates, and times are right-aligned.Boolean types are centered)
                , LEFT = 1 // Left Horizontal Alignment
                , RIGHT = 2 // Right Horizontal Alignment
                , CENTER = 3 // Centered Horizontal Alignment
                , CENTER_CONTINUOUS = 4 // (Center Continuous Horizontal Alignment
                , FILL = 5 // Fill
                , JUSTIFY = 6 // Justify
                , DISTRIBUTED = 7 // Distributed Horizontal Alignment
                ;

        private static final String[] _names = {"general" ,"left" ,"right" ,"center" ,"centerContinuous" ,"fill" ,"justify" ,"distributed"};
        public static int valueOf(String name) {
            return StringUtil.indexOf(_names, name);
        }

        public static String of(int n) {
            return _names[n];
        }
    }
}

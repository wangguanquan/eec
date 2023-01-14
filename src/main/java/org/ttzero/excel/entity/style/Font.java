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

package org.ttzero.excel.entity.style;

import org.dom4j.Element;
import org.ttzero.excel.util.StringUtil;

import java.awt.Color;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.stream.Collectors;

import static org.ttzero.excel.entity.style.Styles.getAttr;

/**
 * @author guanquan.wang at 2018-02-02 16:51
 */
public class Font implements Cloneable {
    private int style;
    // this value is actually the result of the actual font size*10, eg: font-size:11 save as 110
    private int size;
    private String name;
    private Color color;
    private String scheme;
    private int charset;
    private int family;

    private Font() {}

    public Font(String name, int size) {
        this(name, size, Style.NORMAL, null);
    }

    public Font(String name, int size, Color color) {
        this(name, size, Style.NORMAL, color);
    }

    public Font(String name, int size, int style, Color color) {
        this.style = style;
        this.size = checkAndCrop(size * 10);
        this.name = name;
        this.color = color;
    }

    public Font(String name, double size) {
        this(name, size, Style.NORMAL, null);
    }

    public Font(String name, double size, Color color) {
        this(name, size, Style.NORMAL, color);
    }

    public Font(String name, double size, int style, Color color) {
        this.style = style;
        this.size = checkAndCrop(round10(size));
        this.name = name;
        this.color = color;
    }

    /**
     * Create a Font from font-string
     * italic_bold_underLine_size_family_color or italic bold underLine size family color
     * eq: italic_bold_12_宋体 // 斜体 加粗 12号字 宋体
     * eq: bold underLine 12 'Times New Roman' red  // 加粗 12号字 Times New Roman字体 红字
     * @param fontString italic_bold_underLine_size_family_color or italic bold underLine size family color
     * @return the {@link Font}
     */
    public static Font parse(String fontString) throws FontParseException {
        if (fontString.isEmpty()) {
            throw new NullPointerException("Font string empty");
        }
        String s = fontString.trim();
        int i1 = s.indexOf('\''), i2;
        if (i1 >= 0) {
            do {
                i2 = s.indexOf('\'', i1 + 1);
                if (i2 == -1) {
                    throw new FontParseException("Miss end char \"'\"");
                }
                String sub = s.substring(i1, i2 + 1)
                        , mark = sub.substring(1, sub.length() - 1).replace(' ', '+');
                s = s.replace(sub, mark);
                i1 = s.indexOf('\'', i2);
            } while (i1 >= 0);
        }
        String[] values;
        if (s.indexOf('_') >= 0) {
            values = s.split("_");
        } else {
            values = s.split(" ");
        }

        Font font = new Font();
        // The size and family must exist at the same time and the position is unchanged
        boolean beforeSize = true;
        for (int i = 0; i < values.length; i++) {
            String temp = values[i].trim(), v;
            Double size = null;
            if (beforeSize) {
                try {
                    size = Double.valueOf(temp);
                } catch (NumberFormatException e) {
                    //
                }
                if (size == null) {
                    int n;
                    if ((n = temp.indexOf('+')) > 0) {
                        char[] cs = new char[temp.length() - 1];
                        temp.getChars(0, n, cs, 0);
                        temp.getChars(n + 1, temp.length(), cs, n);
                        if (cs[n] >= 'a' && cs[n] <= 'z') {
                            cs[n] -= 32;
                        }
                        v = new String(cs);
                    } else {
                        v = temp;
                    }
                    try {
                        font.style |= Style.valueOf(v);
                    } catch (NoSuchFieldException | IllegalAccessException e) {
                        throw new FontParseException("Property " + v + " not support.");
                    }
                } else if (size > 0) {
                    font.size = checkAndCrop(round10(size));
                    if (i + 1 < values.length) {
                        font.name = values[++i].trim().replace('+', ' ');
                    } else {
                        throw new FontParseException("Font family must after size.");
                    }
                    beforeSize = false;
                } else {
                    throw new FontParseException("Font size must be greater than zero.");
                }
            } else {
                if (temp.indexOf('#') == 0) {
                    font.color = Color.decode(temp);
                } else {
                    try {
                        Field field = Color.class.getDeclaredField(temp);
                        font.color = (Color) field.get(null);
                    } catch (NoSuchFieldException | IllegalAccessException e) {
                        throw new ColorParseException("Color \"" + temp + "\" not support.");
                    }
                }
            }
        }

        return font;
    }

    public int getSize() {
        return size / 10;
    }

    public double getSize2() {
        return size / 10.0D;
    }

    public Font setSize(int size) {
        this.size = checkAndCrop(size * 10);
        return this;
    }

    public Font setSize(double size) {
        this.size = checkAndCrop(round10(size));
        return this;
    }

    public String getName() {
        return name;
    }

    public Font setName(String name) {
        this.name = name;
        return this;
    }

    public int getFamily() {
        return family;
    }

    public Font setFamily(int family) {
        this.family = family;
        return this;
    }

    public Color getColor() {
        return color;
    }

    public Font setColor(Color color) {
        this.color = color;
        return this;
    }

    public int getStyle() {
        return style;
    }

    public Font setStyle(int style) {
        this.style = style;
        return this;
    }

    public String getScheme() {
        return scheme;
    }

    public Font setScheme(String scheme) {
        this.scheme = scheme;
        return this;
    }

    public int getCharset() {
        return charset;
    }

    public Font setCharset(int charset) {
        this.charset = charset;
        return this;
    }

    public Font italic() {
        style |= Style.ITALIC;
        return this;
    }

    public Font bold() {
        style |= Style.BOLD;
        return this;
    }

    public Font underLine() {
        style |= Style.UNDERLINE;
        return this;
    }

    public Font deleteLine() {
        style |= Style.DELETE;
        return this;
    }

    public boolean isItalic() {
        return (style & Style.ITALIC) == Style.ITALIC;
    }
    public boolean isBold() {
        return (style & Style.BOLD) == Style.BOLD;
    }
    public boolean isUnderLine() {
        return (style & Style.UNDERLINE) == Style.UNDERLINE;
    }

    public boolean isDeleteLine() {
        return (style & Style.DELETE) == Style.DELETE;
    }

    public Font delItalic() {
        style &= (Style.UNDERLINE | Style.BOLD | Style.DELETE);
        return this;
    }

    public Font delBold() {
        style &= (Style.UNDERLINE | Style.ITALIC | Style.DELETE);
        return this;
    }

    public Font delUnderLine() {
        style &= (Style.BOLD | Style.ITALIC | Style.DELETE);
        return this;
    }

    public Font delDeleteLine() {
        style &= (Style.UNDERLINE | Style.BOLD | Style.ITALIC);
        return this;
    }

    private static final String[] NODE_NAME = {"u", "b", "i", "strike"};
    @Override
    public String toString() {
        StringBuilder buf = new StringBuilder("<font>");
        // Font style
        for (int n = style, i = 0; n > 0; n >>= 1, i++) {
            if ((n & 1) == 1) buf.append("<").append(NODE_NAME[i]).append("/>");
        }
        // size
        buf.append("<sz val=\"");
        if ((size & 1) == 0) buf.append(size / 10);
        else buf.append(size / 10.0D);
        buf.append("\"/>");
        // color
        if (color != null) {
            int index;
            if ((index = ColorIndex.indexOf(color.getRGB())) == -1) {
                buf.append("<color rgb=\"").append(ColorIndex.toARGB(color.getRGB())).append("\"/>");
            } else {
                buf.append("<color indexed=\"").append(index).append("\"/>");
            }
        }
        // name
        buf.append("<name val=\"").append(name).append("\"/>");
        // family
//        DECORATIVE
//        MODERN
//        NOT_APPLICABLE
//        ROMAN
//        SCRIPT
//        SWISS

        // charset
        if (charset > 0) {
            buf.append("<charset val=\"").append(charset).append("\"/>");
        }
        if (StringUtil.isNotEmpty(scheme)) {
            buf.append("<scheme val=\"").append(scheme).append("\"/>");
        }

        return buf.append("</font>").toString();
    }

    @Override
    public int hashCode() {
        int hash;
        hash = style << 24;
        hash += size << 16;
        hash += name.hashCode() << 8;
        hash += color != null ? color.hashCode() : 0;
        return hash;
    }

    @Override
    public boolean equals(Object o) {
        if (o instanceof Font) {
            Font other = (Font) o;
            return other.family == family
                && other.style == style
                && other.size == size
                && (Objects.equals(other.color, color))
                && (Objects.equals(other.name, name))
                ;
        }
        return false;
    }

    public Element toDom4j(Element root) {
        Element element = root.addElement(StringUtil.lowFirstKey(getClass().getSimpleName()));
        element.addElement("sz").addAttribute("val", ((size & 1) == 0) ? String.valueOf(size / 10) : String.valueOf(size / 10.0D));
        element.addElement("name").addAttribute("val", name);
        if (color != null) {
            int index;
            if ((index = ColorIndex.indexOf(color)) > -1) {
                element.addElement("color").addAttribute("indexed", String.valueOf(index));
            } else {
                element.addElement("color").addAttribute("rgb", ColorIndex.toARGB(color));
            }
        }
        for (int n = style, i = 0; n > 0; n >>= 1, i++) {
            if ((n & 1) == 1) element.addElement(NODE_NAME[i]);
        }

        if (family > 0) {
            element.addElement("family").addAttribute("val", String.valueOf(family));
        }
        if (StringUtil.isNotEmpty(scheme)) {
            element.addElement("scheme").addAttribute("val", scheme);
        }
        if (charset > 0) {
            element.addElement("charset").addAttribute("val", String.valueOf(charset));
        }
        return element;
    }

    public static List<Font> domToFont(Element root) {
        // Fonts tags
        Element ele = root.element("fonts");
        // Break if there don't contains 'fonts' tag
        if (ele == null) {
            return new ArrayList<>();
        }
        return ele.elements().stream().map(Font::parseFontTag).collect(Collectors.toList());
    }

    static Font parseFontTag(Element tag) {
        List<Element> sub = tag.elements();
        Font font = new Font();
        for (Element e : sub) {
            switch (e.getName()) {
                case "sz"     : font.size = round10(Double.parseDouble(getAttr(e, "val"))); break;
                case "color"  : font.color = Styles.parseColor(e);                          break;
                case "name"   : font.name = getAttr(e, "val");                              break;
                case "charset": font.charset = Integer.parseInt(getAttr(e, "val"));         break;
                case "scheme" : font.scheme = getAttr(e, "val");                            break;
                case "family" : font.family = Integer.parseInt(getAttr(e, "val"));          break;
                case "u"      : font.style |= Style.UNDERLINE;                              break;
                case "b"      : font.style |= Style.BOLD;                                   break;
                case "i"      : font.style |= Style.ITALIC;                                 break;
                case "strike" : font.style |= Style.DELETE;                                 break;
            }
        }

        return font;
    }

    @Override public Font clone() {
        Font other;
        try {
            other = (Font) super.clone();
        } catch (CloneNotSupportedException e) {
            other = new Font();
            other.family = family;
            other.charset = charset;
            other.name = name;
            other.scheme = scheme;
        }
        if (color != null) {
            other.color = new Color(color.getRGB());
        }
        return other;
    }

    public static int round10(double v) {
        int i = (int) v;
        double l = v - i;
        if (l < 0.23D) i = i * 10;
        else if (l < 0.73) i = i * 10 + 5;
        else i = i * 10 + 10;
        return i;
    }

    // Check and crop the font size
    static int checkAndCrop(int size) {
        if (size < 10) size = 10;
        else if (size > 4090) size = 4090;
        return size;
    }

    // ######################################Static inner class######################################

    public static class Style {
        public static final int NORMAL    = 0;
        public static final int UNDERLINE = 1;
        public static final int BOLD      = 1 << 1;
        public static final int ITALIC    = 1 << 2;
        public static final int DELETE    = 1 << 3;

        public static int valueOf(String name) throws NoSuchFieldException, IllegalAccessException {
            Field field = Style.class.getDeclaredField(name.toUpperCase());
            return field.getInt(null);
        }
    }

}

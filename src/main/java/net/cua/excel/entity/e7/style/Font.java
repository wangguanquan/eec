package net.cua.excel.entity.e7.style;

import com.sun.istack.internal.NotNull;
import net.cua.excel.manager.Const;
import net.cua.excel.util.StringUtil;
import org.dom4j.Element;

import java.awt.Color;
import java.lang.reflect.Field;

/**
 * Created by guanquan.wang at 2018-02-02 16:51
 */
public class Font {
    private int style;
    private int size;
    private String name;
    private Color color;
    private String scheme;
    private int charset;
    private int family;

    private Font() {}

    public Font(String name, int size) {
        this(name, size, Style.normal, null);
    }

    public Font(String name, int size, Color color) {
        this(name, size, Style.normal, color);
    }

    public Font(String name, int size, int style, Color color) {
        this.style = style;
        this.size = size;
        this.name = name;
        this.color = color;
    }

    /**
     * 字体大小和family必须 颜色默认黑色
     * italic_bold_underLine_size_family_color or italic bold underLine size family color
     * eq: italic_bold_12_宋体 // 斜体 加粗 12号字 宋体
     * eq: bold underLine 12 'Times New Roman' red  // 加粗 12号字 Times New Roman字体 红字
     * @param fontString italic_bold_underLine_size_family_color or italic bold underLine size family color
     * @return
     */
    public static Font parse(@NotNull String fontString) throws FontParseException {
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
        // size family 必须同时存在并且位置不变
        boolean beforeSize = true;
        for (int i = 0; i < values.length; i++) {
            String temp = values[i].trim(), v;
            Integer size = null;
            if (beforeSize) {
                try {
                    size = Integer.valueOf(temp);
                } catch (NumberFormatException e) {
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
                    font.size = size.intValue();
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
        return size;
    }

    public Font setSize(int size) {
        this.size = size;
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
        style |= Style.italic;
        return this;
    }

    public Font bold() {
        style |= Style.bold;
        return this;
    }

    public Font underLine() {
        style |= Style.underLine;
        return this;
    }

    public boolean isItalic() {
        return (style & Style.italic) == Style.italic;
    }
    public boolean isBold() {
        return (style & Style.bold) == Style.bold;
    }
    public boolean isUnderLine() {
        return (style & Style.underLine) == Style.underLine;
    }

    public Font delItalic() {
        style &= (Style.underLine | Style.bold);
        return this;
    }

    public Font delBold() {
        style &= (Style.underLine | Style.italic);
        return this;
    }

    public Font delUnderLine() {
        style &= (Style.bold | Style.italic);
        return this;
    }

    @Override
    public String toString() {
        StringBuilder buf = new StringBuilder("<font>").append(Const.lineSeparator);
        // size
        buf.append("    <sz val=\"").append(size).append("\"/>").append(Const.lineSeparator);
        // color
        if (color != null) {
            int index;
            if ((index = ColorIndex.indexOf(color.getRGB())) == -1) {
                buf.append("    <color rgb=\"").append(ColorIndex.toARGB(color.getRGB())).append("\"/>").append(Const.lineSeparator);
            } else {
                buf.append("    <color indexed=\"").append(index).append("\"/>").append(Const.lineSeparator);
            }
        }
        // name
        buf.append("    <name val=\"").append(name).append("\"/>").append(Const.lineSeparator);
        // family
//        DECORATIVE  装饰
//        MODERN   现代
//        NOT_APPLICABLE  不适用
//        ROMAN
//        SCRIPT
//        SWISS
        switch (style) {
            case 1:
                buf.append("    <u/>").append(Const.lineSeparator);
                break;
            case 2:
                buf.append("    <b/>").append(Const.lineSeparator);
                break;
            case 4:
                buf.append("    <i/>").append(Const.lineSeparator);
                break;
            case 3:
                buf.append("    <u/>").append(Const.lineSeparator);
                buf.append("    <b/>").append(Const.lineSeparator);
                break;
            case 5:
                buf.append("    <i/>").append(Const.lineSeparator);
                buf.append("    <u/>").append(Const.lineSeparator);
                break;
            case 6:
                buf.append("    <b/>").append(Const.lineSeparator);
                buf.append("    <i/>").append(Const.lineSeparator);
                break;
            case 7:
                buf.append("    <i/>").append(Const.lineSeparator);
                buf.append("    <b/>").append(Const.lineSeparator);
                buf.append("    <u/>").append(Const.lineSeparator);
                default:
        }
        // charset
        if (charset > 0) {
            buf.append("    <charset val=\"").append(charset).append("\"/>").append(Const.lineSeparator);
        }
        if (StringUtil.isNotEmpty(scheme)) {
            buf.append("    <scheme val=\"").append(scheme).append("\"/>").append(Const.lineSeparator);
        }

        return buf.append("</font>").toString();
    }

    @Override
    public int hashCode() {
        int hash;
        hash = style << 24;
        hash += size << 16;
        hash += name.hashCode() << 8;
        hash += color.hashCode();
        return hash;
    }

    @Override
    public boolean equals(Object o) {
        if (o instanceof Font) {
            Font other = (Font) o;
            return other.style == style
                    && other.size == size
                    && (other.color != null ? other.color.equals(color) : other.color == color)
                    && (other.name != null ? other.name.equals(name) : other.name == name);
        }
        return false;
    }

    public Element toDom4j(Element root) {
        Element element = root.addElement(StringUtil.lowFirstKey(getClass().getSimpleName()));
        element.addElement("sz").addAttribute("val", String.valueOf(size));
        element.addElement("name").addAttribute("val", name);
        if (color != null) {
            int index;
            if ((index = ColorIndex.indexOf(color)) > -1) {
                element.addElement("color").addAttribute("indexed", String.valueOf(index));
            } else {
                element.addElement("color").addAttribute("rgb", ColorIndex.toARGB(color));
            }
        }
        if (isBold()) {
            element.addElement("b");
        }
        if (isItalic()) {
            element.addElement("i");
        }
        if (isUnderLine()) {
            element.addElement("u");
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

    // ######################################Static inner class######################################

    public static class Style {
        public static final int normal = 0 // 正常
                , italic = 1 << 2 // 斜体
                , bold = 1 << 1 // 加粗
                , underLine = 1 << 0 // 下划线
                ;

        public static final int valueOf(String name) throws NoSuchFieldException, IllegalAccessException {
            Field field = Style.class.getDeclaredField(name);
            return field.getInt(null);
        }
    }

}

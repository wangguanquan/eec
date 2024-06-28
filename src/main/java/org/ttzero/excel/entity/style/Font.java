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
 * 字体，为了简化本字体包含颜色属性，在输出时包含颜色的字体将以指定颜色渲染，
 * 注意：字体在全局样式中是共享的，如果要修改某个属性必须先调用{@link #clone}方法
 * 复制一个字体然后再修改属性，这样才不会影响之前使用过此字体的文本
 *
 * @author guanquan.wang at 2018-02-02 16:51
 */
public class Font implements Cloneable {
    /**
     * 加粗，斜体等字体样式被定义在{@link Font.Style}类中
     */
    private int style;
    /**
     * 字体大小，为发兼容Excel小数字体实际保存的值为乘{@code 10}后的结果，在
     * 例Excel字体大小为{@code 10.5}那么{@code size}为 {@code 105}，
     * 使用{@link #getSize}方法将会得到去尾数的结果{@code 10}，要得到原始值
     * 需要调用{@link #getSize2}
     */
    private int size;
    /**
     * 字体名
     */
    private String name;
    /**
     * 输本文本颜色
     */
    private Color color;
    /**
     * 定义此字体所属的字体方案，有三种取值{@code "none"},{@code "major"}和{@code "minor"},
     * 通常主要字体用于标题等样式，次要字体用于正文和段落文本。
     */
    private String scheme;
    /**
     * 一个整数值，该值指定字体使用的字符集。枚举值参考{@link Charset}，
     * 本类并不会根据字体名自动判断字符集请自行设置
     */
    private int charset;
    /**
     * 字体家族，取值0-255
     *
     * <blockquote><pre>
     * 示例
     *  Value| Font Family
     * ----- +------------
     *     0 | Not applicable.
     *     1 | Roman
     *     2 | Swiss
     *     3 | Modern
     *     4 | Script
     *     5 | Decorative
     * </pre></blockquote>
     */
    private int family;
    /**
     * FontMetrics计算文本宽度
     */
    private transient java.awt.FontMetrics fm;

    private Font() { }

    public Font(String name, int size) {
        this(name, size, Style.PLAIN, null);
    }

    public Font(String name, int size, Color color) {
        this(name, size, Style.PLAIN, color);
    }

    public Font(String name, int size, int style, Color color) {
        this.style = style;
        this.size = checkAndCrop(size * 10);
        this.name = name;
        this.color = color;
        // 这里仅简单判断是否为双字节并设置简体中文
        if (name != null && !name.isEmpty() && name.charAt(0) > 0x4E) this.charset = Charset.GB2312;
    }

    public Font(String name, double size) {
        this(name, size, Style.PLAIN, null);
    }

    public Font(String name, double size, Color color) {
        this(name, size, Style.PLAIN, color);
    }

    public Font(String name, double size, int style, Color color) {
        this.style = style;
        this.size = checkAndCrop(round10(size));
        this.name = name;
        this.color = color;
        // 这里仅简单判断是否为双字节并设置简体中文
        if (name != null && !name.isEmpty() && name.charAt(0) > 0x4E) this.charset = Charset.GB2312;
    }

    /**
     * 解析字符串为字体
     * <p>
     * italic_bold_underline_size_family_color or italic bold underline size family color
     * eq: italic_bold_12_宋体 // 斜体 加粗 12号字 宋体
     * eq: bold underline 12 'Times New Roman' red  // 加粗 12号字 Times New Roman字体 红字
     *
     * @param fontString italic_bold_underline_size_family_color or italic bold underline size family color
     * @return the {@link Font}
     * @throws IllegalArgumentException if convert failed.
     */
    public static Font parse(String fontString) {
        if (fontString.isEmpty()) {
            throw new NullPointerException("Font string empty");
        }
        String s = fontString.trim();
        int i1 = s.indexOf('\''), i2;
        if (i1 >= 0) {
            do {
                i2 = s.indexOf('\'', i1 + 1);
                if (i2 == -1) {
                    throw new IllegalArgumentException("Miss end char \"'\"");
                }
                String sub = s.substring(i1, i2 + 1), mark = sub.substring(1, sub.length() - 1).replace(' ', '+');
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
                        throw new IllegalArgumentException("Property " + v + " not support.");
                    }
                } else if (size > 0) {
                    font.size = checkAndCrop(round10(size));
                    if (i + 1 < values.length) {
                        font.name = values[++i].trim().replace('+', ' ');
                    } else {
                        throw new IllegalArgumentException("Font family must after size.");
                    }
                    beforeSize = false;
                } else {
                    throw new IllegalArgumentException("Font size must be greater than zero.");
                }
            } else {
                if (temp.indexOf('#') == 0) {
                    font.color = Color.decode(temp);
                } else {
                    try {
                        Field field = Color.class.getDeclaredField(temp);
                        font.color = (Color) field.get(null);
                    } catch (NoSuchFieldException | IllegalAccessException e) {
                        throw new IllegalArgumentException("Color \"" + temp + "\" not support.");
                    }
                }
            }
        }

        return font;
    }

    /**
     * 获取去尾数后的字体大小{@code 10.5}实际返回{@code 10}
     *
     * @return 去尾数的字体大小
     */
    public int getSize() {
        return size / 10;
    }

    /**
     * 获取字体大小
     *
     * @return 字体大小
     */
    public double getSize2() {
        return size / 10.0D;
    }

    /**
     * 设置字体大小
     *
     * <p>注意：字体是全局共享的所以修改属性前需要先复制字体</p>
     *
     * @param size 字体大小
     * @return 当前字体
     */
    public Font setSize(int size) {
        this.size = checkAndCrop(size * 10);
        return this;
    }

    /**
     * 设置字体大小
     *
     * <p>注意：字体是全局共享的所以修改属性前需要先复制字体</p>
     *
     * @param size 字体大小
     * @return 当前字体
     */
    public Font setSize(double size) {
        this.size = checkAndCrop(round10(size));
        return this;
    }

    /**
     * 获取字体名
     *
     * @return 字体名
     */
    public String getName() {
        return name;
    }

    /**
     * 设置字体名
     *
     * <p>注意：字体是全局共享的所以修改属性前需要先复制字体</p>
     *
     * @param name 字体名
     * @return 当前字体
     */
    public Font setName(String name) {
        this.name = name;
        return this;
    }

    /**
     * 获取字体家族
     *
     * @return family取值范围{@code 0-255}
     */
    public int getFamily() {
        return family;
    }

    /**
     * 设置字体家族
     *
     * <p>注意：字体是全局共享的所以修改属性前需要先复制字体</p>
     *
     * @param family 取值范围{@code 0-255}
     * @return 当前字体
     */
    public Font setFamily(int family) {
        this.family = family & 0xFF;
        return this;
    }

    /**
     * 获取字体颜色，未主动设置时返回{@code null}，输出到xml时显示黑色
     *
     * @return 字体颜色
     */
    public Color getColor() {
        return color;
    }

    /**
     * 设置字体颜色
     *
     * <p>注意：字体是全局共享的所以修改属性前需要先复制字体</p>
     *
     * @param color 字体颜色
     * @return 当前字体
     */
    public Font setColor(Color color) {
        this.color = color;
        return this;
    }

    /**
     * 获取字体样式，样式定义在{@link Font.Style}类中，建议直接调用专用方法{@link #isBold},
     * {@link #isUnderline}, {@link #isStrikeThru}和{@link #isItalic}方法，
     * 它们直接返回{@code boolean}类型的值方便后续判断
     *
     * @return {@link Font.Style}定义的样式
     */
    public int getStyle() {
        return style;
    }

    /**
     * 设置字体样式，样式定义在{@link Font.Style}类中，建议直接调用专用方法{@link #bold},
     * {@link #underline}, {@link #strikeThru}和{@link #italic}方法设置，
     * 这几个方法可以组合调用最终效果为组合效果
     *
     * <p>注意：字体是全局共享的所以修改属性前需要先复制字体</p>
     *
     * @param style {@link Font.Style}定义的样式
     * @return 当前字体
     */
    public Font setStyle(int style) {
        this.style = style & 31;
        return this;
    }

    /**
     * 获取此字体所属的字体方案，有三种可能的取值{@code "none"},{@code "major"}和{@code "minor"}
     *
     * @return 字体方案
     */
    public String getScheme() {
        return scheme;
    }

    /**
     * 设置此字体所属的字体方案，有三种可能的取值{@code "none"},{@code "major"}和{@code "minor"}
     *
     * <p>注意：字体是全局共享的所以修改属性前需要先复制字体</p>
     *
     * @param scheme {@code "none"},{@code "major"}和{@code "minor"}三种取值之一
     * @return 当前字体
     */
    public Font setScheme(String scheme) {
        if (StringUtil.isNotEmpty(scheme)) {
            scheme = scheme.toLowerCase();
            scheme = ("minor".equals(scheme) || "major".equals(scheme)) ? scheme : null;
        } else scheme = null;
        this.scheme = scheme;
        return this;
    }

    /**
     * 获取字符集
     *
     * @return 字体的字符集参考 {@link Charset}
     */
    public int getCharset() {
        return charset;
    }

    /**
     * 设置字体的字符集
     *
     * <p>注意：字体是全局共享的所以修改属性前需要先复制字体</p>
     *
     * @param charset {@link Charset}
     * @return 当前字体
     */
    public Font setCharset(int charset) {
        this.charset = charset;
        return this;
    }

    /**
     * 添加“斜体”样式
     *
     * <p>注意：字体是全局共享的所以修改属性前需要先复制字体</p>
     *
     * @return 当前字体
     */
    public Font italic() {
        style |= Style.ITALIC;
        return this;
    }

    /**
     * 添加“粗休”样式
     *
     * <p>注意：字体是全局共享的所以修改属性前需要先复制字体</p>
     *
     * @return 当前字体
     */
    public Font bold() {
        style |= Style.BOLD;
        return this;
    }

    /**
     * 添加“下划线”样式
     *
     * <p>注意：字体是全局共享的所以修改属性前需要先复制字体</p>
     *
     * @return 当前字体
     */
    public Font underline() {
        style |= Style.UNDERLINE;
        return this;
    }

    /**
     * 添加“双下划线”样式
     *
     * <p>注意：字体是全局共享的所以修改属性前需要先复制字体</p>
     *
     * @return 当前字体
     */
    public Font doubleUnderline() {
        style |= Style.DOUBLE_UNDERLINE;
        return this;
    }

    /**
     * 添加“删除线”样式
     *
     * <p>注意：字体是全局共享的所以修改属性前需要先复制字体</p>
     *
     * @return 当前字体
     */
    public Font strikeThru() {
        style |= Style.STRIKE;
        return this;
    }

    /**
     * 检查是否有“斜体”样式
     *
     * @return true: 是
     */
    public boolean isItalic() {
        return (style & Style.ITALIC) == Style.ITALIC;
    }

    /**
     * 检查是否有“粗体”样式
     *
     * @return true: 是
     */
    public boolean isBold() {
        return (style & Style.BOLD) == Style.BOLD;
    }

    /**
     * 检查是否有“下划线”样式
     *
     * @return true: 是
     */
    public boolean isUnderline() {
        return (style & Style.UNDERLINE) == Style.UNDERLINE;
    }

    /**
     * 检查是否有“删除线”样式
     *
     * @return true: 是
     */
    public boolean isStrikeThru() {
        return (style & Style.STRIKE) == Style.STRIKE;
    }

    /**
     * 检查是否有“双下划线”样式
     *
     * @return true: 是
     */
    public boolean isDoubleUnderline() {
        return (style & Style.DOUBLE_UNDERLINE) == Style.DOUBLE_UNDERLINE;
    }

    /**
     * 删除"斜体"样式
     *
     * <p>注意：字体是全局共享的所以修改属性前需要先复制字体</p>
     *
     * @return 当前字体
     */
    public Font delItalic() {
        style &= (Style.UNDERLINE | Style.BOLD | Style.STRIKE | Style.DOUBLE_UNDERLINE);
        return this;
    }

    /**
     * 删除"加粗"样式
     *
     * <p>注意：字体是全局共享的所以修改属性前需要先复制字体</p>
     *
     * @return 当前字体
     */
    public Font delBold() {
        style &= (Style.UNDERLINE | Style.ITALIC | Style.STRIKE | Style.DOUBLE_UNDERLINE);
        return this;
    }

    /**
     * 删除"下划线"样式
     *
     * <p>注意：字体是全局共享的所以修改属性前需要先复制字体</p>
     *
     * @return 当前字体
     */
    public Font delUnderline() {
        style &= (Style.BOLD | Style.ITALIC | Style.STRIKE | Style.DOUBLE_UNDERLINE);
        return this;
    }

    /**
     * 删除"下划线"样式
     *
     * <p>注意：字体是全局共享的所以修改属性前需要先复制字体</p>
     *
     * @return 当前字体
     */
    public Font delDoubleUnderline() {
        style &= (Style.BOLD | Style.ITALIC | Style.STRIKE | Style.UNDERLINE);
        return this;
    }

    /**
     * 删除“删除线”样式
     *
     * <p>注意：字体是全局共享的所以修改属性前需要先复制字体</p>
     *
     * @return 当前字体
     */
    public Font delStrikeThru() {
        style &= (Style.UNDERLINE | Style.BOLD | Style.ITALIC | Style.DOUBLE_UNDERLINE);
        return this;
    }

    /**
     * 字体样式对应xml名
     */
    private static final String[] NODE_NAME = {"u", "b", "i", "strike"};

    @Override
    public String toString() {
        StringBuilder buf = new StringBuilder("<font>");
        // Font style
        for (int n = style, i = 0; n > 0; n >>= 1, i++) {
            if ((n & 1) == 1) {
                if (i < NODE_NAME.length) buf.append('<').append(NODE_NAME[i]).append("/>");
                else if (i == 4) buf.append("<u val=\"double\"/>");
            }
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
        if (StringUtil.isNotEmpty(scheme) && !"none".equals(scheme)) {
            buf.append("<scheme val=\"").append(scheme).append("\"/>");
        }

        return buf.append("</font>").toString();
    }

    @Override
    public int hashCode() {
        int hash = size << 16;
        hash += color != null ? color.hashCode() : 0;
        hash += style << 24;
        if (StringUtil.isEmpty(scheme) || "none".equals(scheme)) {
            hash += name.hashCode() << 8;
            hash += charset;
            hash += family;
        } else {
            hash += scheme.hashCode();
        }
        return hash;
    }

    @Override
    public boolean equals(Object o) {
        boolean r = false;
        if (o instanceof Font) {
            Font other = (Font) o;
            r = other.size == size
                && Objects.equals(other.color, color)
                && other.style == style;
            if (r) {
                r = (StringUtil.isEmpty(scheme) || "none".equals(scheme))
                    ? Objects.equals(other.name, name) && other.charset == charset && other.family == family
                    : Objects.equals(other.scheme, scheme);
            }
        }
        return r;
    }

    /**
     * 输出为dom树
     *
     * @param root 父节点
     * @return dom树
     */
    public Element toDom(Element root) {
        Element element = root.addElement(StringUtil.lowFirstKey(getClass().getSimpleName()));
        element.addElement("sz").addAttribute("val", ((size & 1) == 0) ? String.valueOf(size / 10) : String.valueOf(size / 10.0D));
        element.addElement("name").addAttribute("val", name);
        if (color != null) {
            int index;
            if (color instanceof BuildInColor) {
                element.addElement("color").addAttribute("indexed", String.valueOf(((BuildInColor) color).getIndexed()));
            }
            else if ((index = ColorIndex.indexOf(color)) > -1) {
                element.addElement("color").addAttribute("indexed", String.valueOf(index));
            }
            else {
                element.addElement("color").addAttribute("rgb", ColorIndex.toARGB(color));
            }
        }
        for (int n = style, i = 0; n > 0; n >>= 1, i++) {
            if ((n & 1) == 1) {
                if (i < NODE_NAME.length) element.addElement(NODE_NAME[i]);
                else if (i == 4) element.addElement(NODE_NAME[0]).addAttribute("val", "double");
            }
        }

        if (family > 0) {
            element.addElement("family").addAttribute("val", String.valueOf(family));
        }
        if (StringUtil.isNotEmpty(scheme) && !"none".equals(scheme)) {
            element.addElement("scheme").addAttribute("val", scheme);
        }
        if (charset > 0) {
            element.addElement("charset").addAttribute("val", String.valueOf(charset));
        }
        return element;
    }

    /**
     * 将{@code java.awt.Font}字体转为当前字体
     *
     * @param awtFont {@link java.awt.Font}
     * @return a {@code org.ttzero.excel.entity.style.Font}
     */
    public static Font of(java.awt.Font awtFont) {
        return new Font(awtFont.getName(), awtFont.getSize(), awtFont.getStyle() << 1, Color.BLACK);
    }

    /**
     * 将当前字体转为 {@link java.awt.Font}字体
     *
     * @return {@code java.awt.Font}
     */
    public java.awt.Font toAwtFont() {
        return new java.awt.Font(name, style >> 1, getSize());
    }

    /**
     * 通过{@link Font}获取{@code FontMetrics}用以计算文本宽度
     *
     * @return 字体度量对象
     */
    public java.awt.FontMetrics getFontMetrics() {
        return fm != null ? fm : (fm = getFontMetrics(toAwtFont()));
    }

    /**
     * 通过{@link java.awt.Font}获取{@code FontMetrics}用以计算文本宽度
     *
     * @param font awt字体
     * @return 字体度量对象
     */
    public static java.awt.FontMetrics getFontMetrics(java.awt.Font font) {
        return new javax.swing.JLabel().getFontMetrics(font);
    }

    /**
     * 解析字体
     *
     * @param root styles树root
     * @param indexedColors 特殊indexed颜色（大部分情况下为null）
     * @return styles字体
     */
    public static List<Font> domToFont(Element root, Color[] indexedColors) {
        List<Font> fonts = domToFont(root);
        // 替换特殊的indexed颜色
        int indexed;
        for (Font font : fonts) {
            if ((font.color instanceof BuildInColor) && (indexed = ((BuildInColor) font.color).getIndexed()) < indexedColors.length) {
                font.color = indexedColors[indexed];
            }
        }
        return fonts;
    }

    /**
     * 解析字体
     *
     * @param root styles树root
     * @return styles字体
     */
    public static List<Font> domToFont(Element root) {
        // Fonts tags
        Element ele = root.element("fonts");
        // Break if there don't contains 'fonts' tag
        if (ele == null) {
            return new ArrayList<>();
        }
        return ele.elements().stream().map(Font::parseFontTag).collect(Collectors.toList());
    }

    /**
     * 解析xml内容创建字体
     *
     * @param tag dom树font节点
     * @return 字体
     */
    static Font parseFontTag(Element tag) {
        List<Element> sub = tag.elements();
        Font font = new Font();
        for (Element e : sub) {
            switch (e.getName()) {
                case "sz"     : font.size = round10(Double.parseDouble(getAttr(e, "val"))); break;
                case "color"  : font.color = Styles.parseColor(e);                          break;
                case "name"   : font.name = getAttr(e, "val");                              break;
                case "charset": font.charset = Integer.parseInt(getAttr(e, "val"));         break;
                case "scheme" : font.setScheme(getAttr(e, "val"));                          break;
                case "family" : font.family = Integer.parseInt(getAttr(e, "val"));          break;
                case "b"      : font.style |= Style.BOLD;                                   break;
                case "i"      : font.style |= Style.ITALIC;                                 break;
                case "strike" : font.style |= Style.STRIKE;                                 break;
                case "u"      : font.style |= "double".equalsIgnoreCase(e.attributeValue("val")) ? Style.DOUBLE_UNDERLINE : Style.UNDERLINE;
                    break;
            }
        }

        return font;
    }

    @Override
    public Font clone() {
        Font other;
        try {
            other = (Font) super.clone();
        } catch (CloneNotSupportedException e) {
            other = new Font();
            other.family = family;
            other.charset = charset;
            other.name = name;
            other.scheme = scheme;
            other.size = size;
            other.style = style;
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

    /**
     * 字体样式
     */
    public static class Style {
        /**
         * 默认文本
         */
        public static final int PLAIN = 0;
        /**
         * 下划线
         */
        public static final int UNDERLINE = 1;
        /**
         * 粗体
         */
        public static final int BOLD = 1 << 1;
        /**
         * 斜体
         */
        public static final int ITALIC = 1 << 2;
        /**
         * 删除线
         */
        public static final int STRIKE = 1 << 3;
        /**
         * 双下划线
         */
        public static final int DOUBLE_UNDERLINE = 1 << 4;

        public static int valueOf(String name) throws NoSuchFieldException, IllegalAccessException {
            Field field = Style.class.getDeclaredField(name.toUpperCase());
            return field.getInt(null);
        }
    }

}

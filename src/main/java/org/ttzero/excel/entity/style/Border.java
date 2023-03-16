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

import org.ttzero.excel.util.StringUtil;
import org.dom4j.Element;

import java.awt.Color;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Collectors;

import static org.ttzero.excel.entity.style.Styles.getAttr;

/**
 * Border line styles and colors
 *
 * The border contains orientation, style and color. The Excel
 * border has 6 orientations, which are left, right, top, bottom,
 * diagonalDown, diagonalUp. For the style, please refer to {@link BorderStyle}.
 * The default border color is #333333 in EEC
 *
 * @see BorderStyle
 * @author guanquan.wang on 2018-02-06 08:55
 */
public class Border implements Cloneable {

    private static final Color defaultColor = new Color(51, 51, 51); // #333333

    private final SubBorder[] borders;

    public Border() {
        borders = new SubBorder[6]; // left-right-top-bottom-diagonalDown-diagonalUp
    }

    public Border(BorderStyle style, Color color) {
        borders = new SubBorder[6];
        setBorder(style, color);
    }

    /**
     * Top line style
     *
     * @param style the {@link BorderStyle}
     * @return the {@link Border}
     */
    public Border setBorderTop(BorderStyle style) {
        borders[2] = new SubBorder(style, defaultColor);
        return this;
    }

    /**
     * Right line style
     *
     * @param style the {@link BorderStyle}
     * @return the {@link Border}
     */
    public Border setBorderRight(BorderStyle style) {
        borders[1] = new SubBorder(style, defaultColor);
        return this;
    }

    /**
     * Bottom line style
     *
     * @param style the {@link BorderStyle}
     * @return the {@link Border}
     */
    public Border setBorderBottom(BorderStyle style) {
        borders[3] = new SubBorder(style, defaultColor);
        return this;
    }

    /**
     * Left line style
     *
     * @param style the {@link BorderStyle}
     * @return the {@link Border}
     */
    public Border setBorderLeft(BorderStyle style) {
        borders[0] = new SubBorder(style, defaultColor);
        return this;
    }

    /**
     * Top-left to bottom-right line style [\]
     *
     * @param style the {@link BorderStyle}
     * @return the {@link Border}
     */
    public Border setDiagonalDown(BorderStyle style) {
        borders[4] = new SubBorder(style, defaultColor);
        return this;
    }

    /**
     * Bottom-left to top-right line style [/]
     *
     * @param style the {@link BorderStyle}
     * @return the {@link Border}
     */
    public Border setDiagonalUp(BorderStyle style) {
        borders[5] = new SubBorder(style, defaultColor);
        return this;
    }

    /**
     * Setting top-right-bottom-left line style
     *
     * @param style the {@link BorderStyle}
     * @return the {@link Border}
     */
    public Border setBorder(BorderStyle style) {
        borders[0] = new SubBorder(style, defaultColor);
        borders[1] = borders[2] = borders[3] = borders[0];
        return this;
    }


    /**
     * Setting top-left to bottom-right and bottom-left to top-right line style [X]
     *
     * @param style the {@link BorderStyle}
     * @return the {@link Border}
     */
    public Border setDiagonal(BorderStyle style) {
        borders[4] = new SubBorder(style, defaultColor);
        borders[5] = borders[4];
        return this;
    }

    /**
     * Top line style
     *
     * @param style the {@link BorderStyle}
     * @param color the border {@link Color}
     * @return the {@link Border}
     */
    public Border setBorderTop(BorderStyle style, Color color) {
        borders[2] = new SubBorder(style, color);
        return this;
    }

    /**
     * Right line style
     *
     * @param style the {@link BorderStyle}
     * @param color the border {@link Color}
     * @return the {@link Border}
     */
    public Border setBorderRight(BorderStyle style, Color color) {
        borders[1] = new SubBorder(style, color);
        return this;
    }

    /**
     * Bottom line style
     *
     * @param style the {@link BorderStyle}
     * @param color the border {@link Color}
     * @return the {@link Border}
     */
    public Border setBorderBottom(BorderStyle style, Color color) {
        borders[3] = new SubBorder(style, color);
        return this;
    }

    /**
     * Left line style
     *
     * @param style the {@link BorderStyle}
     * @param color the border {@link Color}
     * @return the {@link Border}
     */
    public Border setBorderLeft(BorderStyle style, Color color) {
        borders[0] = new SubBorder(style, color);
        return this;
    }

    /**
     * Top-left to bottom-right line style [\]
     *
     * @param style the {@link BorderStyle}
     * @param color the border {@link Color}
     * @return the {@link Border}
     */
    public Border setDiagonalDown(BorderStyle style, Color color) {
        borders[4] = new SubBorder(style, color);
        return this;
    }

    /**
     * Bottom-left to top-right line style [/]
     *
     * @param style the {@link BorderStyle}
     * @param color the border {@link Color}
     * @return the {@link Border}
     */
    public Border setDiagonalUp(BorderStyle style, Color color) {
        borders[5] = new SubBorder(style, color);
        return this;
    }

    /**
     * Setting top-left to bottom-right and bottom-left to top-right line style [X]
     *
     * @param style the {@link BorderStyle}
     * @param color the border {@link Color}
     * @return the {@link Border}
     */
    public Border setDiagonal(BorderStyle style, Color color) {
        borders[4] = new SubBorder(style, color);
        borders[5] = borders[4];
        return this;
    }

    /**
     * Setting top-right-bottom-left line style
     *
     * @param style the {@link BorderStyle}
     * @param color the border {@link Color}
     * @return the {@link Border}
     */
    public Border setBorder(BorderStyle style, Color color) {
        borders[0] = new SubBorder(style, color);
        borders[1] = borders[0];
        borders[2] = borders[0];
        borders[3] = borders[0];
        return this;
    }

    Border setBorder(int index, BorderStyle style) {
        borders[index] = new SubBorder(style, defaultColor);
        return this;
    }

    Border setBorder(int index, BorderStyle style, Color color) {
        borders[index] = new SubBorder(style, color);
        return this;
    }

    /**
     * Returns top line style
     *
     * @return {@link SubBorder}
     */
    public SubBorder getBorderTop() {
        return borders[2];
    }

    /**
     * Returns right line style
     *
     * @return {@link SubBorder}
     */
    public SubBorder getBorderRight() {
        return borders[1];
    }

    /**
     * Returns bottom line style
     *
     * @return {@link SubBorder}
     */
    public SubBorder getBorderBottom() {
        return borders[3];
    }

    /**
     * Returns left line style
     *
     * @return {@link SubBorder}
     */
    public SubBorder getBorderLeft() {
        return borders[0];
    }

    /**
     * Returns Diagonal-Down line style
     *
     * @return {@link SubBorder}
     */
    public SubBorder getDiagonalDown() {
        return borders[4];
    }

    /**
     * Returns Diagonal-Up line style
     *
     * @return {@link SubBorder}
     */
    public SubBorder getDiagonalUp() {
        return borders[5];
    }

    /**
     * Returns the specify axis border
     *
     * @param axis border axis (left-right-top-bottom-diagonalDown-diagonalUp)
     * @return {@link SubBorder}
     */
    public SubBorder getBorder(int axis) {
        return axis >= 0 && axis < borders.length ? borders[axis] : null;
    }

    /**
     * Returns all borders
     *
     * @return {@link SubBorder}
     */
    public SubBorder[] getBorders() {
        return borders;
    }

    Border delBorder(int index) {
        borders[index] = null;
        return this;
    }

    @Override
    public int hashCode() {
        int down = borders[4] != null ? 1 : 0
            , up = borders[5] != null ? 2 : 0;
        int hash = down | up;
        for (SubBorder sub : borders) {
            hash += sub != null ? sub.hashCode() : 0;
        }
        return hash;
    }

    @Override
    public boolean equals(Object o) {
        if (o instanceof Border) {
            Border other = (Border) o;
            for (int i = 0; i < borders.length; i++) {
                if (other.borders[i] != null) {
                    if (!other.borders[i].equals(borders[i]))
                        return false;
                } else if (borders[i] != null) {
                    return false;
                }
            }
            return true;
        }
        return false;
    }

    /**
     * The setting order is top -&gt; right -&gt; bottom -&gt; left, the
     * attribute order is style-name + color, if the orientation setting
     * is not complete, the unset orientation has no border. If only the
     * orientation is not set, the last color will be complemented.
     * <p>
     * eq:
     * <p>thin red</p>
     * <p>thin red thin dashed dashed</p>
     * <p>medium black thick #cccccc double black hair green</p>
     * <p>none none thin thin</p>
     *
     * @param text the border value
     * @return the parse value of {@link Border}
     */
    public static Border parse(String text) {
        Border border = new Border();
        if (StringUtil.isEmpty(text)) return border;
        String[] values = text.split(" ");
        int index = 0;
        Color color = null;
        for (int i = 0; i < values.length; i++) {
            BorderStyle style = BorderStyle.getByName(values[i]);
            if (style == null) {
                throw new BorderParseException("Border style error.");
            }
            int n = i + 1;
            if (values.length <= n) break;
            String v = values[n];
            BorderStyle style1 = BorderStyle.getByName(v);
            if (style1 == null) {
                if (v.charAt(0) == '#') {
                    color = Color.decode(v);
                } else {
                    try {
                        Field field = Color.class.getDeclaredField(v);
                        color = (Color) field.get(null);
                    } catch (NoSuchFieldException | IllegalAccessException e) {
                        throw new ColorParseException("Color \"" + v + "\" not support.");
                    }
                }
                border.setBorder(index++, style, color);
                i++;
            } else if (color != null) {
                border.setBorder(index++, style, color);
            } else {
                border.setBorder(index++, style);
            }
        }
        if (index == 1) {
            border.borders[1] = border.borders[0];
            border.borders[2] = border.borders[0];
            border.borders[3] = border.borders[0];
        }
        return border;
    }

    public static class SubBorder {
        public final BorderStyle style;
        public final Color color;

        public SubBorder(BorderStyle style, Color color) {
            this.style = style;
            this.color = color;
        }

        public BorderStyle getStyle() {
            return style;
        }

        public Color getColor() {
            return color;
        }

        @Override
        public int hashCode() {
            int hash = color != null ? color.hashCode() : 0;
            return (style.ordinal() << 24) | (hash << 8 >>> 8);
        }

        @Override
        public boolean equals(Object o) {
            return (o instanceof SubBorder) && o.hashCode() == hashCode();
        }
    }

    static final String[] direction = {"left", "right", "top", "bottom", "diagonal", "diagonal"};

    public Element toDom4j(Element root) {
        Element element = root.addElement(StringUtil.lowFirstKey(getClass().getSimpleName()));
        for (int i = 0; i < direction.length; i++) {
            Element sub = element.element(direction[i]);
            if (sub == null) sub = element.addElement(direction[i]);
            writeProperties(sub, borders[i]);
        }

        boolean down = borders[4] != null, up = borders[5] != null;
        if (down) {
            element.addAttribute("diagonalDown", "1");
        }
        if (up) {
            element.addAttribute("diagonalUp", "1");
        }
        return element;
    }

    public static List<Border> domToBorder(Element root) {
        // Borders tags
        Element ele = root.element("borders");
        // Break if there don't contains 'borders' tag
        if (ele == null) {
            return new ArrayList<>();
        }
        return ele.elements().stream().map(Border::parseBorderTag).collect(Collectors.toList());
    }

    static Border parseBorderTag(Element tag) {
        List<Element> sub = tag.elements();
        // Diagonal attr
        String diagonalDown = getAttr(tag, "diagonalDown");
        int padding = ("1".equals(diagonalDown) || "true".equalsIgnoreCase(diagonalDown)) ? 1 : 0;
        String diagonalUp = getAttr(tag, "diagonalUp");
        padding |= ("1".equals(diagonalUp) || "true".equalsIgnoreCase(diagonalUp)) ? 1 << 1 : 0;

        Border border = new Border();
        for (Element e : sub) {
            int i = StringUtil.indexOf(direction, e.getName());
            // unknown element
            if (i < 0) continue;
            BorderStyle style = BorderStyle.getByName(getAttr(e, "style"));
            if (style == null) style = BorderStyle.NONE;
            Color color = Styles.parseColor(e.element("color"));
            if (i < 4) border.setBorder(i, style, color);
            else if ((padding & 1) == 1) border.setBorder(4, style, color);
            else if ((padding & 2) == 2) border.setBorder(5, style, color);
        }

        return border;
    }

    protected void writeProperties(Element element, SubBorder subBorder) {
        if (subBorder != null && subBorder.style != BorderStyle.NONE) {
            element.addAttribute("style", subBorder.style.getName());
            Element colorEle = element.element("color");
            if (colorEle == null) colorEle = element.addElement("color");
            int index;
            if (subBorder.color instanceof BuildInColor) {
                colorEle.addAttribute("indexed", String.valueOf(((BuildInColor) subBorder.color).getIndexed()));
            }
            else if ((index = ColorIndex.indexOf(subBorder.color)) > -1) {
                element.addAttribute("indexed", String.valueOf(index));
            }
            else {
                colorEle.addAttribute("rgb", ColorIndex.toARGB(subBorder.color));
            }
        }
    }

    @Override
    public String toString() {
        return (borders[0] != null ? borders[0].style : BorderStyle.NONE) + " "
            + (borders[1] != null ? borders[1].style : BorderStyle.NONE) + " "
            + (borders[2] != null ? borders[2].style : BorderStyle.NONE) + " "
            + (borders[3] != null ? borders[3].style : BorderStyle.NONE) + " "
            + (borders[4] != null ? borders[4].style : BorderStyle.NONE) + " "
            + (borders[5] != null ? borders[5].style : BorderStyle.NONE);
    }

    @Override
    public Border clone() {
        Border newBorder = new Border();
        for (int i = 0; i < borders.length; i++) {
            if (borders[i] == null) continue;
            newBorder.borders[i] = new SubBorder(borders[i].style, borders[i].color);
        }
        return newBorder;
    }
}

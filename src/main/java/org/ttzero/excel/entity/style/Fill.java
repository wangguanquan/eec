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
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.stream.Collectors;

import static org.ttzero.excel.entity.style.Styles.getAttr;

/**
 * 填充，在样式中位于第{@code 18-24}位，目前只支持单色填充
 *
 * @author guanquan.wang at 2018-02-06 08:55
 */
public class Fill implements Cloneable {
    private PatternType patternType;
    private Color bgColor, fgColor;

    public Fill(PatternType patternType, Color bgColor, Color fgColor) {
        this.patternType = patternType;
        this.bgColor = bgColor;
        this.fgColor = fgColor;
    }

    public Fill(PatternType patternType, Color fgColor) {
        this.patternType = patternType;
        this.fgColor = fgColor;
    }

    public Fill(PatternType patternType) {
        this.patternType = patternType;
    }

    public Fill(Color fgColor) {
        this.fgColor = fgColor;
        this.patternType = PatternType.solid;
    }

    public Fill() {}

    public PatternType getPatternType() {
        return patternType;
    }

    public Fill setPatternType(PatternType patternType) {
        this.patternType = patternType;
        return this;
    }

    public Color getBgColor() {
        return bgColor;
    }

    public Fill setBgColor(Color bgColor) {
        this.bgColor = bgColor;
        return this;
    }

    public Color getFgColor() {
        return fgColor;
    }

    public Fill setFgColor(Color fgColor) {
        this.fgColor = fgColor;
        return this;
    }

    @Override
    public int hashCode() {
        int hash = 0;
        if (patternType != null) {
            hash += patternType.ordinal() << 24;
        }
        int c = 0;
        if (bgColor != null) {
            c += bgColor.hashCode();
        }
        if (fgColor != null) {
            c += fgColor.hashCode();
        }
        return hash + (c << 8 >>> 8);
    }

    @Override
    public boolean equals(Object o) {
        if (o instanceof Fill) {
            Fill other = (Fill) o;
            return (other.patternType == this.patternType)
                && (Objects.equals(other.bgColor, bgColor))
                && (Objects.equals(other.fgColor, fgColor));
        }
        return false;
    }

    /**
     * fgColor bgColor patternType
     * color patternType : fgColor patternType
     *
     * @param text the Fill string value
     * @return the parse value of {@link Fill}
     */
    public static Fill parse(String text) {
        Fill fill = new Fill();
        if (StringUtil.isNotEmpty(text)) {
            String[] values = text.split(" ");
            for (String v : values) {
                PatternType patternType;
                try {
                    patternType = PatternType.valueOf(v);
                } catch (IllegalArgumentException e) {
                    patternType = null;
                }
                if (patternType == null) {
                    Color color = Styles.toColor(v);
                    if (fill.fgColor == null) {
                        fill.fgColor = color;
                    } else {
                        fill.bgColor = color;
                    }
                } else {
                    fill.patternType = patternType;
                }
            }
        }
//        if (fill.patternType == null) {
//            fill.patternType = PatternType.solid;
//        }
        return fill;
    }

    Element toDom(Element root) {
        if (patternType == null) {
            patternType = PatternType.solid;
        }
        Element element = root.addElement(StringUtil.lowFirstKey(getClass().getSimpleName()));
        Element patternFill = element.addElement("patternFill").addAttribute("patternType", patternType.name());

        int index;
        if (fgColor != null) {
            if (fgColor instanceof BuildInColor) {
                patternFill.addElement("fgColor").addAttribute("indexed", String.valueOf(((BuildInColor) fgColor).getIndexed()));
            }
            else if ((index = ColorIndex.indexOf(fgColor)) > -1) {
                patternFill.addElement("fgColor").addAttribute("indexed", String.valueOf(index));
            }
            else {
                patternFill.addElement("fgColor").addAttribute("rgb", ColorIndex.toARGB(fgColor));
            }
        }
        if (bgColor != null) {
            if (bgColor instanceof BuildInColor) {
                patternFill.addElement("bgColor").addAttribute("indexed", String.valueOf(((BuildInColor) bgColor).getIndexed()));
            }
            else if ((index = ColorIndex.indexOf(bgColor)) > -1) {
                patternFill.addElement("bgColor").addAttribute("indexed", String.valueOf(index));
            }
            else {
                patternFill.addElement("bgColor").addAttribute("rgb", ColorIndex.toARGB(bgColor));
            }
        }

        return element;
    }

    /**
     * 解析Dom树并转为填充对象
     *
     * @param root dom树
     * @param indexedColors 特殊indexed颜色（大部分情况下为null）
     * @return 填充
     */
    public static List<Fill> domToFill(Element root, Color[] indexedColors) {
        List<Fill> fills = domToFill(root);
        int indexed;
        for (Fill fill : fills) {
            if (fill.fgColor instanceof BuildInColor && (indexed = ((BuildInColor) fill.fgColor).getIndexed()) < indexedColors.length) {
                fill.fgColor = indexedColors[indexed];
            }
            if (fill.bgColor instanceof BuildInColor && (indexed = ((BuildInColor) fill.bgColor).getIndexed()) < indexedColors.length) {
                fill.bgColor = indexedColors[indexed];
            }
        }
        return fills;
    }

    /**
     * 解析Dom树并转为填充对象
     *
     * @param root dom树
     * @return 填充
     */
    public static List<Fill> domToFill(Element root) {
        // Fills tags
        Element ele = root.element("fills");
        // Break if there don't contains 'fills' tag
        if (ele == null) {
            return new ArrayList<>();
        }
        return ele.elements().stream().map(Fill::parseFillTag).collect(Collectors.toList());
    }

    static Fill parseFillTag(Element tag) {
        Fill fill = new Fill();
        // 单色背景
        Element e = tag.element("patternFill");
        if (e != null) {
            try {
                fill.patternType = PatternType.valueOf(getAttr(e, "patternType"));
            } catch (IllegalArgumentException ex) {
                // Ignore
            }
            Element fgColor = e.element("fgColor");
            if (fgColor != null) {
                fill.fgColor = Styles.parseColor(fgColor);
            }
            Element bgColor = e.element("bgColor");
            if (bgColor != null) {
                fill.bgColor = Styles.parseColor(bgColor);
            }
        }
        // FIXME 双色背景目前仅简单支持（取双色中的起始色）
        else if ((e = tag.element("gradientFill")) != null) {
            List<Element> sub = e.elements("stop");
            if (sub != null && !sub.isEmpty()) {
                Element sub0 = sub.get(0).element("color");
                fill.fgColor = Styles.parseColor(sub0);
                fill.patternType = PatternType.solid;
            }
        }
        return fill;
    }

    @Override public Fill clone() {
        Fill other;
        try {
            other = (Fill) super.clone();
        } catch (CloneNotSupportedException e) {
            other = new Fill();
            other.patternType = patternType;
        }
        if (bgColor != null) {
            other.bgColor = new Color(bgColor.getRGB());
        }
        if (fgColor != null) {
            other.fgColor = new Color(fgColor.getRGB());
        }
        return other;
    }

    @Override public String toString() {
        if (patternType == null || patternType == PatternType.none) return "<fill><patternFill patternType=\"none\"/></fill>";
        StringBuilder buf = new StringBuilder("<fill><patternFill patternType=\"").append(patternType).append("\">");
        int index;
        if (fgColor != null) {
            if (fgColor instanceof BuildInColor)
                buf.append("<fgColor indexed=\"").append(((BuildInColor) fgColor).getIndexed()).append("\"/>");
            else if ((index = ColorIndex.indexOf(fgColor)) > -1)
                buf.append("<fgColor indexed=\"").append(index).append("\"/>");
            else buf.append("<fgColor rgb=\"").append(ColorIndex.toARGB(fgColor)).append("\"/>");
        }
        if (bgColor != null) {
            if (bgColor instanceof BuildInColor)
                buf.append("<bgColor indexed=\"").append(((BuildInColor) bgColor).getIndexed()).append("\"/>");
            else if ((index = ColorIndex.indexOf(bgColor)) > -1)
                buf.append("<bgColor indexed=\"").append(index).append("\"/>");
            else buf.append("<bgColor rgb=\"").append(ColorIndex.toARGB(bgColor)).append("\"/>");
        }
        buf.append("</patternFill></fill>");
        return buf.toString();
    }
}

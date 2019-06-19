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

package org.ttzero.excel.entity.style;


import org.dom4j.Element;
import org.ttzero.excel.util.StringUtil;

import java.awt.Color;
import java.lang.reflect.Field;

/**
 * Created by guanquan.wang at 2018-02-06 08:55
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

    public boolean equals(Object o) {
        if (o instanceof Fill) {
            Fill other = (Fill) o;
            return (other.patternType == this.patternType)
                && (other.bgColor != null ? other.bgColor.equals(bgColor) : null == bgColor)
                && (other.fgColor != null ? other.fgColor.equals(fgColor) : null == fgColor);
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
                    Color color;
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
        if (fill.patternType == null) {
            fill.patternType = PatternType.solid;
        }
        return fill;
    }

    Element toDom4j(Element root) {
        if (patternType == null) {
            patternType = PatternType.solid;
        }
        Element element = root.addElement(StringUtil.lowFirstKey(getClass().getSimpleName()));
        Element patternFill = element.addElement("patternFill").addAttribute("patternType", patternType.name());

        if (fgColor != null) {
            int colorIndex = ColorIndex.indexOf(fgColor);
            if (colorIndex > -1) {
                patternFill.addElement("fgColor").addAttribute("indexed", String.valueOf(colorIndex));
            } else {
                patternFill.addElement("fgColor").addAttribute("rgb", ColorIndex.toARGB(fgColor));
            }
        }
        if (bgColor != null) {
            int colorIndex = ColorIndex.indexOf(bgColor);
            if (colorIndex > -1) {
                patternFill.addElement("bgColor").addAttribute("indexed", String.valueOf(colorIndex));
            } else {
                patternFill.addElement("bgColor").addAttribute("rgb", ColorIndex.toARGB(bgColor));
            }
        }

        return element;
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
}

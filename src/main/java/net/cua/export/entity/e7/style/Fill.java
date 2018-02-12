package net.cua.export.entity.e7.style;


import net.cua.export.util.StringUtil;
import org.dom4j.Element;

import java.awt.Color;
import java.lang.reflect.Field;

/**
 * Created by guanquan.wang at 2018-02-06 08:55
 */
public class Fill {
    private PatternType patternType;
    private Color bgColor, fgColor;

    public Fill(PatternType patternType, Color bgColor, Color fgColor) {
        this.patternType = patternType;
        this.bgColor = bgColor;
        this.fgColor = fgColor;
    }

    public Fill(PatternType patternType, Color bgColor) {
        this.patternType = patternType;
        this.bgColor = bgColor;
    }

    public Fill(PatternType patternType) {
        this.patternType = patternType;
    }

    public Fill(Color bgColor) {
        this.bgColor = bgColor;
        this.patternType = PatternType.solid;
    }

    public Fill() {}

    public PatternType getPatternType() {
        return patternType;
    }

    public void setPatternType(PatternType patternType) {
        this.patternType = patternType;
    }

    public Color getBgColor() {
        return bgColor;
    }

    public void setBgColor(Color bgColor) {
        this.bgColor = bgColor;
    }

    public Color getFgColor() {
        return fgColor;
    }

    public void setFgColor(Color fgColor) {
        this.fgColor = fgColor;
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
                    && (other.bgColor != null ? other.bgColor.equals(bgColor) : other.bgColor == bgColor)
                    && (other.fgColor != null ? other.fgColor.equals(fgColor) : other.fgColor == fgColor);
        }
        return false;
    }

    /**
     * fgColor bgColor patternType
     * color patternType : fgColor patternType
     * @param text
     * @return
     */
    public static final Fill parse(String text) {
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

    public Element toDom4j(Element root) {
        if (patternType == null) {
            patternType = PatternType.solid;
        }
        Element element = root.addElement(StringUtil.lowFirstKey(getClass().getSimpleName()));
        Element patternFill = element.addElement("patternFill").addAttribute("patternType", patternType.name());

        if (bgColor != null) {
            int colorIndex = ColorIndex.indexOf(bgColor);
            if (colorIndex > -1) {
                patternFill.addElement("bgColor").addAttribute("indexed", String.valueOf(colorIndex));
            } else {
                patternFill.addElement("bgColor").addAttribute("rgb", ColorIndex.toARGB(bgColor));
            }
        }
        if (fgColor != null) {
            int colorIndex = ColorIndex.indexOf(fgColor);
            if (colorIndex > -1) {
                patternFill.addElement("fgColor").addAttribute("indexed", String.valueOf(colorIndex));
            } else {
                patternFill.addElement("fgColor").addAttribute("rgb", ColorIndex.toARGB(fgColor));
            }
        }
        return element;
    }
}

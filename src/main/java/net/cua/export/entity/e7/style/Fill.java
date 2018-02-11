package net.cua.export.entity.e7.style;


import net.cua.export.util.StringUtil;
import org.dom4j.Element;

import java.awt.*;
import java.lang.reflect.Field;

/**
 * Created by wanggq at 2018-02-06 08:55
 */
public class Fill {
//    public final static short     NO_FILL             = 0  ;
//    public final static short     SOLID_FILL          = 1  ;
//    public final static short     FINE_DOTS           = 2  ;
//    public final static short     ALT_BARS            = 3  ;
//    public final static short     SPARSE_DOTS         = 4  ;
//    public final static short     THICK_HORZ_BANDS    = 5  ;
//    public final static short     THICK_VERT_BANDS    = 6  ;
//    public final static short     THICK_BACKWARD_DIAG = 7  ;
//    public final static short     THICK_FORWARD_DIAG  = 8  ;
//    public final static short     BIG_SPOTS           = 9  ;
//    public final static short     BRICKS              = 10 ;
//    public final static short     THIN_HORZ_BANDS     = 11 ;
//    public final static short     THIN_VERT_BANDS     = 12 ;
//    public final static short     THIN_BACKWARD_DIAG  = 13 ;
//    public final static short     THIN_FORWARD_DIAG   = 14 ;
//    public final static short     SQUARES             = 15 ;
//    public final static short     DIAMONDS            = 16 ;

    private PatternType patternType;
    private Color bgColor, fgColor;

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

    public boolean equals(Object o) {
        if (o instanceof Fill) {
            Fill other = (Fill) o;
            return other.patternType == patternType && other.bgColor == bgColor && other.fgColor == fgColor;
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
        if (patternType == null || patternType == PatternType.none) {
            return null;
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

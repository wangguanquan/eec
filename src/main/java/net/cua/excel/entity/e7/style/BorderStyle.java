package net.cua.excel.entity.e7.style;

/**
 * 定义边框样式
 */
public enum BorderStyle {
    NONE("none"),
    THIN("thin"),    // 细直线
    MEDIUM("medium"),
    DASHED("dashed"), // 虚线
    DOTTED("dotted"), // 点线
    THICK("thick"), // 粗直线
    DOUBLE("double"), // 双线
    HAIR("hair"),
    MEDIUM_DASHED("mediumDashed"),
    DASH_DOT("dashDot"),
    MEDIUM_DASH_DOT("mediumDashDot"),
    DASH_DOT_DOT("dashDotDot"),
    MEDIUM_DASH_DOT_DOT("mediumDashDotDot"),
    SLANTED_DASH_DOT("slantDashDot");

    private String name;

    BorderStyle(String name) {
        this.name = name;
    }

    public String getName() {
        return name;
    }

    public static BorderStyle getByName(String name) {
        BorderStyle[] borderStyles = values();
        for (BorderStyle borderStyle : borderStyles) {
            if (borderStyle.name.equals(name)) {
                return borderStyle;
            }
        }
        return null;
    }
}

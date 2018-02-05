package net.cua.export.entity.e7.style;

import java.awt.Color;

/**
 * Created by wanggq at 2018-02-02 16:51
 */
public class Font {
    private int fontSize;
    private String fontName;
    private Color color;
    public Font () {

    }

    public Font(int fontSize, String fontName, Color color) {
        this.fontSize = fontSize;
        this.fontName = fontName;
        this.color = color;
    }

    public int getFontSize() {
        return fontSize;
    }

    public void setFontSize(int fontSize) {
        this.fontSize = fontSize;
    }

    public String getFontName() {
        return fontName;
    }

    public void setFontName(String fontName) {
        this.fontName = fontName;
    }

    public Color getColor() {
        return color;
    }

    public void setColor(Color color) {
        this.color = color;
    }
}

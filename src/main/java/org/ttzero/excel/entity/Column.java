/*
 * Copyright (c) 2017-2021, guanquan.wang@yandex.com All Rights Reserved.
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


package org.ttzero.excel.entity;

import org.ttzero.excel.entity.style.Border;
import org.ttzero.excel.entity.style.Fill;
import org.ttzero.excel.entity.style.Font;
import org.ttzero.excel.entity.style.Horizontals;
import org.ttzero.excel.entity.style.NumFmt;
import org.ttzero.excel.entity.style.Styles;
import org.ttzero.excel.entity.style.Verticals;
import org.ttzero.excel.processor.IntConversionProcessor;
import org.ttzero.excel.processor.StyleProcessor;

import static org.ttzero.excel.entity.IWorksheetWriter.isBigDecimal;
import static org.ttzero.excel.entity.IWorksheetWriter.isBool;
import static org.ttzero.excel.entity.IWorksheetWriter.isChar;
import static org.ttzero.excel.entity.IWorksheetWriter.isDate;
import static org.ttzero.excel.entity.IWorksheetWriter.isDateTime;
import static org.ttzero.excel.entity.IWorksheetWriter.isDouble;
import static org.ttzero.excel.entity.IWorksheetWriter.isFloat;
import static org.ttzero.excel.entity.IWorksheetWriter.isInt;
import static org.ttzero.excel.entity.IWorksheetWriter.isLocalDate;
import static org.ttzero.excel.entity.IWorksheetWriter.isLocalDateTime;
import static org.ttzero.excel.entity.IWorksheetWriter.isLocalTime;
import static org.ttzero.excel.entity.IWorksheetWriter.isLong;
import static org.ttzero.excel.entity.IWorksheetWriter.isString;
import static org.ttzero.excel.entity.IWorksheetWriter.isTime;
import static org.ttzero.excel.entity.style.Styles.INDEX_BORDER;

/**
 * Associated with Worksheet for controlling head style and cache
 * column data types and conversions
 *
 * @author guanquan.wang at 2021-08-29 19:49
 */
public class Column {
    /**
     * The key of Map or field name of entry
     */
    public String key;
    /**
     * The header name
     */
    public String name;
    /**
     * The cell type
     */
    public Class<?> clazz;
    /**
     * The string value is shared
     */
    public boolean share;
    /**
     * The int value conversion
     */
    public IntConversionProcessor processor;
    /**
     * The style conversion
     */
    public StyleProcessor styleProcessor;
    /**
     * The style of cell
     */
    public int cellStyle;
    /**
     * The style of header
     */
    public int headerStyle;
    /**
     * The style index of cell, -1 if not be setting
     */
    private int cellStyleIndex = -1;
    /**
     * The style index of header, -1 if not be setting
     */
    private int headerStyleIndex = -1;
    /**
     * The cell width
     */
    public double width;
    public int o;
    public Styles styles;
    public Comment headerComment, cellComment;
    /**
     * Specify the cell number format
     */
    public NumFmt numFmt;
    /**
     * Only export column name and ignore value
     */
    public boolean ignoreValue;
    /**
     * Wrap text in a cell
     */
    public int wrapText;
    /**
     * Specify the column index
     */
    public int colIndex = -1;

    /**
     * Constructor Column
     */
    public Column() { }

    /**
     * Constructor Column
     *
     * @param name  the column name
     * @param clazz the cell type
     */
    public Column(String name, Class<?> clazz) {
        this(name, clazz, true);
    }

    /**
     * Constructor Column
     *
     * @param name the column name
     * @param key  field
     */
    public Column(String name, String key) {
        this(name, key, true);
    }

    /**
     * Constructor Column
     *
     * @param name  the column name
     * @param key   field
     * @param clazz the cell type
     */
    public Column(String name, String key, Class<?> clazz) {
        this(name, key, true);
        this.clazz = clazz;
    }

    /**
     * Constructor Column
     *
     * @param name      the column name
     * @param clazz     the cell type
     * @param processor The int value conversion
     */
    public Column(String name, Class<?> clazz, IntConversionProcessor processor) {
        this(name, clazz, processor, true);
    }

    /**
     * Constructor Column
     *
     * @param name      the column name
     * @param key       field
     * @param processor The int value conversion
     */
    public Column(String name, String key, IntConversionProcessor processor) {
        this(name, key, processor, true);
    }

    /**
     * Constructor Column
     *
     * @param name  the column name
     * @param clazz the cell type
     * @param share true:shared false:inline string
     */
    public Column(String name, Class<?> clazz, boolean share) {
        this.name = name;
        this.clazz = clazz;
        this.share = share;
    }

    /**
     * Constructor Column
     *
     * @param name  the column name
     * @param key   filed
     * @param share true:shared false:inline string
     */
    public Column(String name, String key, boolean share) {
        this.name = name;
        this.key = key;
        this.share = share;
    }

    /**
     * Constructor Column
     *
     * @param name      the column name
     * @param clazz     the cell type
     * @param processor The int value conversion
     * @param share     true:shared false:inline string
     */
    public Column(String name, Class<?> clazz, IntConversionProcessor processor, boolean share) {
        this(name, clazz, share);
        this.processor = processor;
    }

    /**
     * Constructor Column
     *
     * @param name      the column name
     * @param key       field
     * @param clazz     type of cell
     * @param processor The int value conversion
     */
    public Column(String name, String key, Class<?> clazz, IntConversionProcessor processor) {
        this(name, key, clazz);
        this.processor = processor;
    }

    /**
     * Constructor Column
     *
     * @param name      the column name
     * @param key       field
     * @param processor The int value conversion
     * @param share     true:shared false:inline string
     */
    public Column(String name, String key, IntConversionProcessor processor, boolean share) {
        this(name, key, share);
        this.processor = processor;
    }

    /**
     * Constructor Column
     *
     * @param name      the column name
     * @param clazz     the cell type
     * @param cellStyle the style of cell
     */
    public Column(String name, Class<?> clazz, int cellStyle) {
        this(name, clazz, cellStyle, true);
    }

    /**
     * Constructor Column
     *
     * @param name      the column name
     * @param key       field
     * @param cellStyle the style of cell
     */
    public Column(String name, String key, int cellStyle) {
        this(name, key, cellStyle, true);
    }

    /**
     * Constructor Column
     *
     * @param name      the column name
     * @param clazz     the cell type
     * @param cellStyle the style of cell
     * @param share     true:shared false:inline string
     */
    public Column(String name, Class<?> clazz, int cellStyle, boolean share) {
        this(name, clazz, share);
        this.cellStyle = cellStyle;
    }

    /**
     * Constructor Column
     *
     * @param name      the column name
     * @param key       field
     * @param cellStyle the style of cell
     * @param share     true:shared false:inline string
     */
    public Column(String name, String key, int cellStyle, boolean share) {
        this(name, key, share);
        this.cellStyle = cellStyle;
    }

    /**
     * Setting the cell's width
     *
     * @param width the width value
     * @return the {@link Column}
     */
    public Column setWidth(double width) {
        if (width < 0.00000001) {
            throw new ExcelWriteException("Width " + width + " less than 0.");
        }
        this.width = width;
        return this;
    }

    /**
     * Setting the cell is shared
     *
     * @return true:shared false:inline string
     */
    public boolean isShare() {
        return share;
    }

    /**
     * Returns the column name
     *
     * @return the column name
     */
    public String getName() {
        return name;
    }

    /**
     * Setting the column name
     *
     * @param name the column name
     * @return the {@link Column}
     */
    public Column setName(String name) {
        this.name = name;
        return this;
    }

    /**
     * Returns the cell type
     *
     * @return the cell type
     */
    public Class<?> getClazz() {
        return clazz;
    }

    /**
     * Setting the cell type
     *
     * @param clazz the cell type
     * @return the {@link Column}
     */
    public Column setClazz(Class<?> clazz) {
        this.clazz = clazz;
        return this;
    }

    /**
     * Setting the int value conversion
     *
     * @param processor The int value conversion
     * @return the {@link Column}
     */
    public Column setProcessor(IntConversionProcessor processor) {
        this.processor = processor;
        return this;
    }

    /**
     * Setting the style conversion
     *
     * @param styleProcessor The style conversion
     * @return the {@link Column}
     */
    public Column setStyleProcessor(StyleProcessor styleProcessor) {
        this.styleProcessor = styleProcessor;
        return this;
    }

    /**
     * Returns the width of cell
     *
     * @return the cell width
     */
    public double getWidth() {
        return width;
    }

    /**
     * Setting the cell's style
     *
     * @param cellStyle the styles value
     * @return the {@link Column}
     */
    public Column setCellStyle(int cellStyle) {
        this.cellStyle = cellStyle;
        this.cellStyleIndex = styles.of(cellStyle);
        return this;
    }

    /**
     * Setting the header's style
     *
     * @param headerStyle the styles value
     * @return the {@link Column}
     */
    public Column setHeaderStyle(int headerStyle) {
        this.headerStyle = headerStyle;
        this.headerStyleIndex = styles.of(headerStyle);
        return this;
    }

    /**
     * Settings the x-axis of column in row
     *
     * @param colIndex column index (zero base)
     * @return the {@link Column}
     */
    public Column setColIndex(int colIndex) {
        this.colIndex = colIndex;
        return this;
    }

    /**
     * Returns the style index of cell, -1 if not be setting
     *
     * @return index of style
     */
    public int getCellStyleIndex() {
        return cellStyleIndex;
    }

    /**
     * Returns the header style index of cell, -1 if not be setting
     *
     * @return index of style
     */
    public int getHeaderStyleIndex() {
        return headerStyleIndex;
    }

    /**
     * Returns the default horizontal style
     * the Date, Character, Bool has center value, the
     * Numeric has right value, otherwise left value
     *
     * @return the horizontal value
     */
    int defaultHorizontal() {
        int horizontal;
        if (isDate(clazz) || isDateTime(clazz)
                || isLocalDate(clazz) || isLocalDateTime(clazz)
                || isTime(clazz) || isLocalTime(clazz)
                || isChar(clazz) || isBool(clazz)) {
            horizontal = Horizontals.CENTER;
        } else if (isInt(clazz) || isLong(clazz)
                || isFloat(clazz) || isDouble(clazz)
                || isBigDecimal(clazz)) {
            horizontal = Horizontals.RIGHT;
        } else {
            horizontal = Horizontals.LEFT;
        }
        return horizontal;
    }

    /**
     * Setting the cell styles
     *
     * @param font the font
     * @return the {@link Column}
     */
    public Column setCellStyle(Font font) {
        this.cellStyle = styles.of(
                (font != null ? styles.addFont(font) : 0)
                        | Verticals.CENTER
                        | defaultHorizontal());
        return this;
    }

    /**
     * Setting the cell styles
     *
     * @param font       the font
     * @param horizontal the horizontal style
     * @return the {@link Column}
     */
    public Column setCellStyle(Font font, int horizontal) {
        this.cellStyle = styles.of(
                (font != null ? styles.addFont(font) : 0)
                        | Verticals.CENTER
                        | horizontal);
        return this;
    }

    /**
     * Setting the cell styles
     *
     * @param font   the font
     * @param border the border style
     * @return the {@link Column}
     */
    public Column setCellStyle(Font font, Border border) {
        this.cellStyle = styles.of(
                (font != null ? styles.addFont(font) : 0)
                        | (border != null ? styles.addBorder(border) : 0)
                        | Verticals.CENTER
                        | defaultHorizontal());
        return this;
    }

    /**
     * Setting the cell styles
     *
     * @param font       the font
     * @param border     the border style
     * @param horizontal the horizontal style
     * @return the {@link Column}
     */
    public Column setCellStyle(Font font, Border border, int horizontal) {
        this.cellStyle = styles.of(
                (font != null ? styles.addFont(font) : 0)
                        | (border != null ? styles.addBorder(border) : 0)
                        | Verticals.CENTER
                        | horizontal);
        return this;
    }

    /**
     * Setting the cell styles
     *
     * @param font   the font
     * @param fill   the fill style
     * @param border the border style
     * @return the {@link Column}
     */
    public Column setCellStyle(Font font, Fill fill, Border border) {
        this.cellStyle = styles.of(
                (font != null ? styles.addFont(font) : 0)
                        | (fill != null ? styles.addFill(fill) : 0)
                        | (border != null ? styles.addBorder(border) : 0)
                        | Verticals.CENTER
                        | defaultHorizontal());
        return this;
    }

    /**
     * Setting the cell styles
     *
     * @param font       the font
     * @param fill       the fill style
     * @param border     the border style
     * @param horizontal the horizontal style
     * @return the {@link Column}
     */
    public Column setCellStyle(Font font, Fill fill, Border border, int horizontal) {
        this.cellStyle = styles.of(
                (font != null ? styles.addFont(font) : 0)
                        | (fill != null ? styles.addFill(fill) : 0)
                        | (border != null ? styles.addBorder(border) : 0)
                        | Verticals.CENTER
                        | horizontal);
        return this;
    }

    /**
     * Setting the cell styles
     *
     * @param font       the font
     * @param fill       the fill style
     * @param border     the border style
     * @param vertical   the vertical style
     * @param horizontal the horizontal style
     * @return the {@link Column}
     */
    public Column setCellStyle(Font font, Fill fill, Border border, int vertical, int horizontal) {
        this.cellStyle = styles.of(
                (font != null ? styles.addFont(font) : 0)
                        | (fill != null ? styles.addFill(fill) : 0)
                        | (border != null ? styles.addBorder(border) : 0)
                        | vertical
                        | horizontal);
        return this;
    }

    /**
     * Setting the cell styles
     *
     * @param numFmt     the number format
     * @param font       the font
     * @param fill       the fill style
     * @param border     the border style
     * @param vertical   the vertical style
     * @param horizontal the horizontal style
     * @return the {@link Column}
     */
    public Column setCellStyle(NumFmt numFmt, Font font, Fill fill, Border border, int vertical, int horizontal) {
        this.cellStyle = styles.of(
                (numFmt != null ? styles.addNumFmt(numFmt) : 0)
                        | (font != null ? styles.addFont(font) : 0)
                        | (fill != null ? styles.addFill(fill) : 0)
                        | (border != null ? styles.addBorder(border) : 0)
                        | vertical
                        | horizontal);
        return this;
    }

    /**
     * Setting cell string value is shared
     * Shared is only valid for strings, and the converted cell type
     * is also valid for strings.
     *
     * @param share true:shared false:inline string
     * @return the {@link Column}
     */
    public Column setShare(boolean share) {
        this.share = share;
        return this;
    }

    /**
     * Setting a cell format of number or date type
     *
     * @param code the format string
     * @return the {@link Column}
     */
    public Column setNumFmt(String code) {
        this.numFmt = new NumFmt(code);
        return this;
    }

    /**
     * Returns the column {@link NumFmt}
     *
     * @return number format
     */
    public NumFmt getNumFmt() {
        return numFmt != null ? numFmt : styles.getNumFmt(cellStyle);
    }

    /**
     * Returns default style based on cell type
     *
     * @param clazz the cell type
     * @return the styles value
     */
    public int getCellStyle(Class<?> clazz) {
        int style;
        if (isString(clazz)) {
            style = Styles.defaultStringBorderStyle() | wrapText;
        } else if (isDateTime(clazz) || isLocalDateTime(clazz)) {
            style = styles.addNumFmt(new NumFmt("yyyy\\-mm\\-dd\\ hh:mm:ss")) | (1 << INDEX_BORDER) | Horizontals.CENTER;
        } else if (isDate(clazz) || isLocalDate(clazz)) {
            style = styles.addNumFmt(new NumFmt("yyyy\\-mm\\-dd")) | (1 << INDEX_BORDER) | Horizontals.CENTER;
        } else if (isBool(clazz) || isChar(clazz)) {
            style = Styles.clearHorizontal(Styles.defaultStringBorderStyle()) | Horizontals.CENTER;
        } else if (isInt(clazz) || isLong(clazz)) {
            style = Styles.defaultIntBorderStyle();
        } else if (isFloat(clazz) || isDouble(clazz) || isBigDecimal(clazz)) {
            style = Styles.defaultDoubleBorderStyle();
        } else if (isTime(clazz) || isLocalTime(clazz)) {
            style =  styles.addNumFmt(new NumFmt("hh:mm:ss")) | (1 << INDEX_BORDER) | Horizontals.CENTER;
        } else {
            style = (1 << Styles.INDEX_FONT) | (1 << INDEX_BORDER); // Auto-style
        }

        // Reset custom number format if specified.
        if (numFmt != null) {
            style = Styles.clearNumFmt(style) | styles.addNumFmt(numFmt);
        }

        return style;
    }

    /**
     * Setting the cell styles
     *
     * @return the styles value
     */
    public int getCellStyle() {
        if (cellStyleIndex != -1) {
            return cellStyle;
        }
        setCellStyle(getCellStyle(clazz));
        return cellStyle;
    }

    /**
     * Ignore value
     *
     * @return the {@link Column} self
     */
    public Column ignoreValue() {
        this.ignoreValue = true;
        return this;
    }

    /**
     * Wrap text in a cell
     * <p>
     * Microsoft Excel can wrap text so it appears on multiple lines in a cell.
     * You can format the cell so the text wraps automatically, or enter a manual line break.
     *
     * @param wrapText set wrap
     * @return the {@link Column} self
     */
    public Column setWrapText(boolean wrapText) {
        this.wrapText = wrapText ? 1 : 0;
        return this;
    }
}

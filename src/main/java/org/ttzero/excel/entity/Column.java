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
import org.ttzero.excel.manager.Const;
import org.ttzero.excel.processor.ConversionProcessor;
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
import static org.ttzero.excel.entity.style.NumFmt.DATETIME_FORMAT;
import static org.ttzero.excel.entity.style.NumFmt.DATE_FORMAT;
import static org.ttzero.excel.entity.style.NumFmt.TIME_FORMAT;

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
    public ConversionProcessor processor;
    /**
     * The style conversion
     */
    public StyleProcessor styleProcessor;
    /**
     * The style of cell
     */
    public Integer cellStyle;
    /**
     * The style of header
     */
    public Integer headerStyle;
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
    public double o;
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
     * The previous Column point
     *
     * Support multi header columns
     *
     * @since 0.5.1
     */
    public Column prev;
    /**
     * The next Column point
     *
     * Support multi header columns
     *
     * @since 0.5.1
     */
    public Column next;
    /**
     * The tail Column point
     *
     * Support multi header columns
     *
     * @since 0.5.1
     */
    public Column tail;
    /**
     * The real col-Index used to write
     */
    int realColIndex;

    /**
     * Constructor Column
     */
    public Column() { }

    /**
     * Constructor Column
     *
     * @param name  the column name
     */
    public Column(String name) {
        this.name = name;
    }

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
    public Column(String name, Class<?> clazz, ConversionProcessor processor) {
        this(name, clazz, processor, true);
    }

    /**
     * Constructor Column
     *
     * @param name      the column name
     * @param key       field
     * @param processor The int value conversion
     */
    public Column(String name, String key, ConversionProcessor processor) {
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
    public Column(String name, Class<?> clazz, ConversionProcessor processor, boolean share) {
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
    public Column(String name, String key, Class<?> clazz, ConversionProcessor processor) {
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
    public Column(String name, String key, ConversionProcessor processor, boolean share) {
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
    public Column setProcessor(ConversionProcessor processor) {
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
        if (styles != null) this.cellStyleIndex = styles.of(cellStyle);
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
        if (styles != null) this.headerStyleIndex = styles.of(headerStyle);
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
        return cellStyleIndex >= 0 ? cellStyleIndex : (cellStyleIndex = styles != null && cellStyle != null ? styles.of(cellStyle) : -1);
    }

    /**
     * Returns the header style index of cell, -1 if not be setting
     *
     * @return index of style
     */
    public int getHeaderStyleIndex() {
        return headerStyleIndex >= 0 ? headerStyleIndex : (headerStyleIndex = styles != null && headerStyle != null ? styles.of(headerStyle) : -1);
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
     * Setting a cell format of number or date type
     *
     * @param numFmt {@link NumFmt}
     * @return the {@link Column}
     */
    public Column setNumFmt(NumFmt numFmt) {
        this.numFmt = numFmt;
        return this;
    }

    /**
     * Returns the column {@link NumFmt}
     *
     * @return number format
     */
    public NumFmt getNumFmt() {
        return numFmt != null ? numFmt : (numFmt = styles.getNumFmt(cellStyle));
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
        } else if (isDateTime(clazz) || isDate(clazz) || isLocalDateTime(clazz)) {
            if (numFmt == null) numFmt = DATETIME_FORMAT;
            style = (1 << INDEX_BORDER) | Horizontals.CENTER;
        } else if (isBool(clazz) || isChar(clazz)) {
            style = Styles.clearHorizontal(Styles.defaultStringBorderStyle()) | Horizontals.CENTER;
        } else if (isInt(clazz) || isLong(clazz)) {
            style = Styles.defaultIntBorderStyle();
        } else if (isFloat(clazz) || isDouble(clazz) || isBigDecimal(clazz)) {
            style = Styles.defaultDoubleBorderStyle();
        } else if (isLocalDate(clazz)) {
            if (numFmt == null) numFmt = DATE_FORMAT;
            style = (1 << INDEX_BORDER) | Horizontals.CENTER;
        } else if (isTime(clazz) || isLocalTime(clazz)) {
            if (numFmt == null) numFmt = TIME_FORMAT;
            style =  (1 << INDEX_BORDER) | Horizontals.CENTER;
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
        if (cellStyle != null) {
            return cellStyle;
        }
        setCellStyle(getCellStyle(clazz));
        return cellStyle;
    }

    /**
     * @return bool
     */
    public boolean isIgnoreValue() {
        return ignoreValue;
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

    /**
     * Setting the header cell comment
     *
     * @param headerComment {@link Comment}
     * @return the {@link Column} self
     */
    public Column setHeaderComment(Comment headerComment) {
        this.headerComment = headerComment;
        return this;
    }

    /**
     * Append sub-column at the tail
     *
     * @param column a sub-column
     * @return the {@link Column} self
     */
    public Column addSubColumn(Column column) {
        if (this == column) {
            return this;
        }
        if (next != null) {
            int subSize = subColumnSize(), appendSize = column.subColumnSize();
            if (subSize + appendSize > Const.Limit.HEADER_SUB_COLUMNS) {
                throw new ExcelWriteException("Too many sub-column occur. Max support " + Const.Limit.HEADER_SUB_COLUMNS + ", current is " + subSize);
            }
            column.prev = this.tail;
            this.tail.next = column;
        } else {
            this.next = column;
            column.prev = this;
        }
        this.tail = column.tail != null ? column.tail : column;
        return this;
    }

    /**
     * Returns the size of sub-column
     *
     * @return size of sub-column(include root column)
     */
    public int subColumnSize() {
        int i = 1;
        if (next != null) {
            Column next = this.next;
            for (; next != tail; next = next.next, i++);
            i++;
        }
        return i;
    }

    /**
     * Returns an array containing all of the sub-column
     *
     * @return an array containing all of the Column
     */
    public Column[] toArray() {
        return toArray(new Column[subColumnSize()]);
    }

    /**
     * Returns an array containing all of the sub-column
     *
     * @param dist the array into which the elements of the column are to be stored
     * @throws NullPointerException if the specified array is null
     */
    public Column[] toArray(Column[] dist) {
        int len = Math.min(subColumnSize(), dist.length);
        if (len < 1) return dist;
        Column e = this;
        for (int i = 0; i < len; i++) {
            dist[i] = e;
            e = e.next;
        }
        return dist;
    }

    /**
     * Returns the real col-index(one base)
     */
    public int getRealColIndex() {
        return realColIndex;
    }
}

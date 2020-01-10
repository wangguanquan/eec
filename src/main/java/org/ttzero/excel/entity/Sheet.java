/*
 * Copyright (c) 2017, guanquan.wang@yandex.com All Rights Reserved.
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


import org.ttzero.excel.annotation.TopNS;
import org.ttzero.excel.entity.e7.XMLWorksheetWriter;
import org.ttzero.excel.entity.style.Border;
import org.ttzero.excel.entity.style.Fill;
import org.ttzero.excel.entity.style.Font;
import org.ttzero.excel.entity.style.Horizontals;
import org.ttzero.excel.entity.style.NumFmt;
import org.ttzero.excel.entity.style.Styles;
import org.ttzero.excel.entity.style.Verticals;
import org.ttzero.excel.manager.Const;
import org.ttzero.excel.manager.RelManager;
import org.ttzero.excel.processor.IntConversionProcessor;
import org.ttzero.excel.processor.StyleProcessor;
import org.ttzero.excel.reader.Cell;
import org.ttzero.excel.util.FileUtil;

import java.awt.Color;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.ObjectInputStream;
import java.io.ObjectOutputStream;
import java.lang.reflect.InvocationTargetException;
import java.nio.file.Path;

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
import static org.ttzero.excel.manager.Const.ROW_BLOCK_SIZE;
import static org.ttzero.excel.util.StringUtil.isEmpty;

/**
 * Each worksheet corresponds to one or more sheet.xml of physical.
 * When the amount of data exceeds the upper limit of the worksheet,
 * the extra data will be written in the next worksheet page of the
 * current position, with the name of the parent worksheet. After
 * adding "(1,2,3...n)" as the name of the copied sheet, the pagination
 * is automatic without additional settings.
 * <p>
 * Usually worksheetWriter calls the
 * {@link #nextBlock} method to load a row-block for writing.
 * When the row-block returns the flag EOF, mean is the current worksheet
 * finished written, and the next worksheet is written.
 * <p>
 * Extends the existing worksheet to implement a custom data source worksheet.
 * The data source can be micro-services, Mybatis, JPA or any others. If
 * the data source returns an array of json objects, please convert to
 * an object ArrayList or Map ArrayList, the object ArrayList needs to
 * extends {@link ListSheet}, the Map ArrayList needs to extends
 * {@link ListMapSheet} and implement the {@link ListSheet#more} method.
 * <p>
 * If other formats cannot be converted to ArrayList, you
 * need to inherit from the base class {@link Sheet} and implement the
 * {@link #resetBlockData} and {@link #getHeaderColumns} methods.
 *
 * @see ListSheet
 * @see ListMapSheet
 * @see ResultSetSheet
 * @see StatementSheet
 * @see CSVSheet
 *
 * @author guanquan.wang on 2017/9/26.
 */
@TopNS(prefix = {"", "r"}, value = "worksheet", uri = {Const.SCHEMA_MAIN, Const.Relationship.RELATIONSHIP})
public abstract class Sheet implements Cloneable, Storageable {
    protected Workbook workbook;

    protected String name;
    protected Column[] columns;
    protected WaterMark waterMark;
    protected RelManager relManager;
    protected int id;
    /**
     * To mark the cell auto-width
     */
    protected int autoSize;
    /**
     * The default cell width
     */
    protected double width = 20;
    /**
     * The row number
     */
    protected int rows;
    /**
     * Mark the cell is hidden
     */
    protected boolean hidden;

    protected int headStyle;
    /**
     * Automatic interlacing color
     */
    protected int autoOdd = -1;
    /**
     * Odd row's background color
     */
    protected int oddFill;
    /**
     * A copy worksheet flag
     */
    protected boolean copySheet;
    protected int copyCount;

    protected RowBlock rowBlock;
    protected IWorksheetWriter sheetWriter;
    /**
     * To mark the header column is ready
     */
    protected boolean headerReady;
    /**
     * Close resource on the last copy worksheet
     */
    protected boolean shouldClose = true;

    protected ICellValueAndStyle cellValueAndStyle;

    public int getId() {
        return id;
    }

    public void setId(int id) {
        this.id = id;
    }

    public void setSheetWriter(IWorksheetWriter sheetWriter) {
        this.sheetWriter = sheetWriter;
    }

    public void setCellValueAndStyle(ICellValueAndStyle cellValueAndStyle) {
        this.cellValueAndStyle = cellValueAndStyle;
    }

    public Sheet() {
        this(null);
    }

    /**
     * Constructor worksheet
     *
     * @param name the worksheet name
     */
    public Sheet(String name) {
        this.name = name;
        relManager = new RelManager();
    }

    /**
     * Constructor worksheet
     *
     * @param name    the worksheet name
     * @param columns the header info
     */
    public Sheet(String name, final Column... columns) {
        this(name, null, columns);
    }

    /**
     * Constructor worksheet
     *
     * @param name      the worksheet name
     * @param waterMark the water mark
     * @param columns   the header info
     */
    public Sheet(String name, WaterMark waterMark, final Column... columns) {
        this.name = name;
        this.columns = columns;
        this.waterMark = waterMark;
        relManager = new RelManager();
    }

    /**
     * Associated with Worksheet for controlling head style and cache
     * column data types and conversions
     */
    public static class Column {
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
        public boolean share = true;
        /**
         * 0: standard 1:percentage 2:RMB
         */
        public int type;
        /**
         * The int value conversion
         */
        public IntConversionProcessor processor;
        /**
         * The style conversion
         */
        public StyleProcessor styleProcessor;
        /**
         * The style of cell, -1 if not be setting
         */
        public int cellStyle = -1;
        /**
         * The cell width
         */
        public double width;
        public int o;
        public Styles styles;

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
         * @return the {@link Sheet.Column}
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
         * Setting the cell type
         *
         * @param type the cell type
         * @return the {@link Sheet.Column}
         */
        public Column setType(int type) {
            this.type = type;
            return this;
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
         * @return the {@link Sheet.Column}
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
         * @return the {@link Sheet.Column}
         */
        public Column setClazz(Class<?> clazz) {
            this.clazz = clazz;
            return this;
        }

        /**
         * Setting the int value conversion
         *
         * @param processor The int value conversion
         * @return the {@link Sheet.Column}
         */
        public Column setProcessor(IntConversionProcessor processor) {
            this.processor = processor;
            return this;
        }

        /**
         * Setting the style conversion
         *
         * @param styleProcessor The style conversion
         * @return the {@link Sheet.Column}
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
         * @return the {@link Sheet.Column}
         */
        public Column setCellStyle(int cellStyle) {
            this.cellStyle = cellStyle;
            return this;
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
         * @return the {@link Sheet.Column}
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
         * @return the {@link Sheet.Column}
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
         * @return the {@link Sheet.Column}
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
         * @return the {@link Sheet.Column}
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
         * @return the {@link Sheet.Column}
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
         * @return the {@link Sheet.Column}
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
         * @return the {@link Sheet.Column}
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
         * @return the {@link Sheet.Column}
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
         * @return the {@link Sheet.Column}
         */
        public Column setShare(boolean share) {
            this.share = share;
            return this;
        }

        private static final NumFmt ip = new NumFmt("0%_);[Red]\\(0%\\)") // 整数百分比
            , ir = new NumFmt("¥0_);[Red]\\(¥0\\)") // 整数人民币
            , fp = new NumFmt("0.00%_);[Red]\\(0.00%\\)") // 小数百分比
            , fr = new NumFmt("¥0.00_);[Red]\\(¥0.00\\)") // 小数人民币
            , tm = new NumFmt("hh:mm:ss") // 时分秒
            ;

        /**
         * Returns default style based on cell type
         *
         * @param clazz the cell type
         * @return the styles value
         */
        public int getCellStyle(Class<?> clazz) {
            int style;
            if (isString(clazz)) {
                style = Styles.defaultStringBorderStyle();
            } else if (isDate(clazz) || isLocalDate(clazz)) {
                style = Styles.defaultDateBorderStyle();
            } else if (isDateTime(clazz) || isLocalDateTime(clazz)) {
                style = Styles.defaultTimestampBorderStyle();
            } else if (isBool(clazz) || isChar(clazz)) {
                style = Styles.clearHorizontal(Styles.defaultStringBorderStyle()) | Horizontals.CENTER;
            } else if (isInt(clazz) || isLong(clazz)) {
                style = Styles.defaultIntBorderStyle();
                switch (type) {
                    case Const.ColumnType.NORMAL: // 正常显示数字
                        break;
                    case Const.ColumnType.PARENTAGE: // 百分比显示
                        style = Styles.clearNumfmt(style) | styles.addNumFmt(ip);
                        break;
                    case Const.ColumnType.RMB: // 显示人民币
                        style = Styles.clearNumfmt(style) | styles.addNumFmt(ir);
                        break;
                    default:
                }
            } else if (isFloat(clazz) || isDouble(clazz) || isBigDecimal(clazz)) {
                style = Styles.defaultDoubleBorderStyle();
                switch (type) {
                    case Const.ColumnType.NORMAL: // 正常显示数字
                        break;
                    case Const.ColumnType.PARENTAGE: // 百分比显示
                        style = Styles.clearNumfmt(style) | styles.addNumFmt(fp);
                        break;
                    case Const.ColumnType.RMB: // 显示人民币
                        style = Styles.clearNumfmt(style) | styles.addNumFmt(fr);
                        break;
                    default:
                }
            } else if (isTime(clazz) || isLocalTime(clazz)) {
                style = Styles.clearNumfmt(Styles.defaultDateBorderStyle()) | styles.addNumFmt(tm);
            } else {
                style = 0; // Auto-style
            }
            return style;
        }

        /**
         * Setting the cell styles
         *
         * @return the styles value
         */
        public int getCellStyle() {
            if (cellStyle != -1) {
                return cellStyle;
            }
            return cellStyle = getCellStyle(clazz);
        }
    }

    /**
     * Returns workbook
     *
     * @return the {@link Workbook}
     */
    public Workbook getWorkbook() {
        return workbook;
    }

    /**
     * Setting the workbook
     *
     * @param workbook the {@link Workbook}
     * @return the {@link Sheet}
     */
    public Sheet setWorkbook(Workbook workbook) {
        this.workbook = workbook;
        if (columns != null) {
            for (int i = 0; i < columns.length; i++) {
                columns[i].styles = workbook.getStyles();
            }
        }
        return this;
    }

    /**
     * Output the export detail info
     *
     * @param code the message code in message properties file
     */
    public void what(String code) {
        workbook.what(code);
    }

    /**
     * Output export detail info
     *
     * @param code the message code in message properties file
     * @param args the placeholder values
     */
    public void what(String code, String... args) {
        workbook.what(code, args);
    }

    /**
     * Returns shared string
     *
     * @return the {@link SharedStrings}
     */
    public SharedStrings getSst() {
        return workbook.getSst();
    }

    /**
     * Return the cell default width
     *
     * @return the width value
     */
    public double getDefaultWidth() {
        return width;
    }

    /**
     * Setting auto resize cell's width
     *
     * @return the {@link Sheet}
     */
    public Sheet autoSize() {
        this.autoSize = 1;
        return this;
    }

    /**
     * Setting fix column width
     *
     * @return the {@link Sheet}
     */
    public Sheet fixSize() {
        this.autoSize = 2;
        return this;
    }

    /**
     * Setting fix column width
     *
     * @param width the column width
     * @return the {@link Sheet}
     */
    public Sheet fixSize(double width) {
        this.autoSize = 2;
        this.width = width;
        if (headerReady) {
            for (Column hc : columns) {
                hc.setWidth(width);
            }
        }
        return this;
    }

    /**
     * Returns the re-size setting
     *
     * @return 1: auto-size 2:fix-size
     */
    public int getAutoSize() {
        return autoSize;
    }

    /**
     * Test is auto size column width
     *
     * @return true if auto-size
     */
    public boolean isAutoSize() {
        return autoSize == 1;
    }

    /**
     * Cancel the odd row's fill style
     *
     * @return the {@link Sheet}
     */
    public Sheet cancelOddStyle() {
        this.autoOdd = 1;
        return this;
    }

    public int getAutoOdd() {
        return autoOdd;
    }

    public void setAutoOdd(int autoOdd) {
        this.autoOdd = autoOdd;
    }

    /**
     * Setting the odd row's fill style
     *
     * @param fill the fill style
     * @return the {@link Sheet}
     */
    public Sheet setOddFill(Fill fill) {
        this.oddFill = workbook.getStyles().addFill(fill);
        return this;
    }

    /**
     * Returns the odd columns fill style
     *
     * @return the fill style value
     */
    public int getOddFill() {
        return oddFill;
    }

    /**
     * Returns the worksheet name
     *
     * @return the worksheet name
     */
    public String getName() {
        return name;
    }

    /**
     * Setting the worksheet name
     *
     * @param name the worksheet name
     * @return the {@link Sheet}
     */
    public Sheet setName(String name) {
        this.name = name;
        return this;
    }

    /**
     * Returns the header column info
     * <p>
     * The copy sheet will use the parent worksheet header information.
     *
     * @return array of column
     */
    public Column[] getHeaderColumns() {
        if (!headerReady) {
            if (columns == null) {
                columns = new Column[0];
            }
            headerReady = true;
        }
        return columns;
    }

    /**
     * Setting the header rows's columns
     *
     * @param columns the header row's columns
     * @return the {@link Sheet}
     */
    public Sheet setColumns(final Column[] columns) {
        this.columns = columns.clone();
        for (int i = 0; i < columns.length; i++) {
            columns[i].styles = workbook.getStyles();
        }
        return this;
    }

    /**
     * Returns the {@link WaterMark}
     *
     * @return the {@link WaterMark}
     * @see WaterMark
     */
    public WaterMark getWaterMark() {
        return waterMark;
    }

    /**
     * Setting the {@link WaterMark}
     *
     * @param waterMark the {@link WaterMark}
     * @return the {@link Sheet}
     */
    public Sheet setWaterMark(WaterMark waterMark) {
        this.waterMark = waterMark;
        return this;
    }

    /**
     * Returns the worksheet is hidden
     *
     * @return true: hidden, false: not hidden
     */
    public boolean isHidden() {
        return hidden;
    }

    /**
     * Setting the worksheet status
     *
     * @return the {@link Sheet}
     */
    public Sheet hidden() {
        this.hidden = true;
        return this;
    }

    /**
     * abstract method close
     *
     * @throws IOException if I/O error occur
     */
    public void close() throws IOException {
        if (sheetWriter != null) {
            sheetWriter.close();
        }
    }

    /**
     * Write worksheet data to path
     *
     * @param path the storage path
     * @throws IOException if I/O error occur
     */
    @Override
    public void writeTo(Path path) throws IOException {
        if (sheetWriter == null) {
            throw new ExcelWriteException("Worksheet writer is not instanced.");
        }
        if (!copySheet) {
            paging();
        }
        if (!headerReady) {
            getHeaderColumns();
        }
        if (rowBlock == null) {
            rowBlock = new RowBlock(getRowBlockSize());
        } else rowBlock.reopen();

        sheetWriter.writeTo(path);
    }

    /**
     * Split worksheet data
     */
    protected void paging() { }

    /**
     * Add relationship
     *
     * @param rel Relationship
     * @return worksheet
     */
    public Sheet addRel(Relationship rel) {
        relManager.add(rel);
        return this;
    }

    public Relationship find(String key) {
        return relManager.likeByTarget(key);
    }

    /**
     * Returns the worksheet name
     *
     * @return name of worksheet
     */
    public String getFileName() {
        return "sheet" + id + cellValueAndStyle.getFileSuffix();
    }

    /**
     * Setting the header column styles
     *
     * @param font   the font
     * @param fill   the fill style
     * @param border the border style
     * @return the {@link Sheet}
     */
    public Sheet setHeadStyle(Font font, Fill fill, Border border) {
        return setHeadStyle(null, font, fill, border, Verticals.CENTER, Horizontals.CENTER);
    }

    /**
     * Setting the header column styles
     *
     * @param font       the font
     * @param fill       the fill style
     * @param border     the border style
     * @param vertical   the vertical style
     * @param horizontal the horizontal style
     * @return the {@link Sheet}
     */
    public Sheet setHeadStyle(Font font, Fill fill, Border border, int vertical, int horizontal) {
        return setHeadStyle(null, font, fill, border, vertical, horizontal);
    }

    /**
     * Setting the header column styles
     *
     * @param numFmt     the number format
     * @param font       the font
     * @param fill       the fill style
     * @param border     the border style
     * @param vertical   the vertical style
     * @param horizontal the horizontal style
     * @return the {@link Sheet}
     */
    public Sheet setHeadStyle(NumFmt numFmt, Font font, Fill fill, Border border, int vertical, int horizontal) {
        Styles styles = workbook.getStyles();
        headStyle = styles.of(
            (numFmt != null ? styles.addNumFmt(numFmt) : 0)
                | (font != null ? styles.addFont(font) : 0)
                | (fill != null ? styles.addFill(fill) : 0)
                | (border != null ? styles.addBorder(border) : 0)
                | vertical
                | horizontal);
        return this;
    }

    /**
     * Setting the header cell styles
     *
     * @param style the styles value
     * @return the {@link Sheet}
     */
    public Sheet setHeadStyle(int style) {
        headStyle = style;
        return this;
    }

    public int defaultHeadStyle() {
        if (headStyle == 0) {
            Styles styles = workbook.getStyles();
            Font font = new Font(workbook.getI18N().getOrElse("local-font-family", "Arial")
                , 11, Font.Style.bold, Color.white);
            headStyle = styles.of(styles.addFont(font)
                | styles.addFill(Fill.parse("solid #666699"))
                | styles.addBorder(Border.parse("thin black"))
                | Verticals.CENTER
                | Horizontals.CENTER);
        }
        return headStyle;
    }

    protected static boolean nonOrIntDefault(int style) {
        return style == -1
            || style == Styles.defaultIntBorderStyle()
            || style == Styles.defaultIntStyle();
    }

    /**
     * Returns total rows in this worksheet
     *
     * @return -1 if unknown or uncertain
     */
    public int size() {
        return -1;
    }

    /**
     * Returns a row-block. The row-block is content by 32 rows
     *
     * @return a row-block
     */
    public RowBlock nextBlock() {
        // clear first
        rowBlock.clear();

        resetBlockData();

        return rowBlock.flip();
    }

    /**
     * The worksheet is written by units of row-block. The default size
     * of a row-block is 32, which means that 32 rows of data are
     * written at a time. If the data is not enough, the {@code more()}
     * method will be called to get more data.
     *
     * @return the row-block size
     */
    public int getRowBlockSize() {
        return ROW_BLOCK_SIZE;
    }

    /**
     * Write some final info
     *
     * @param workSheetPath the worksheet path
     * @throws IOException if I/O error occur
     */
    public void afterSheetAccess(Path workSheetPath) throws IOException {
        // relationship
        if (sheetWriter instanceof XMLWorksheetWriter) {
            relManager.write(workSheetPath, getFileName());
        }

        // others ...
    }

    /**
     * Returns the copy worksheet name
     *
     * @return the name
     */
    protected String getCopySheetName() {
        int sub = copyCount;
        String _name = name;
        // reset name
        int i = name.lastIndexOf('(');
        if (i > 0) {
            int fs = Integer.parseInt(name.substring(i + 1, name.lastIndexOf(')')));
            _name = name.substring(0, name.charAt(i - 1) == ' ' ? i - 1 : i);
            if (++fs > sub) sub = fs;
        }
        return _name + " (" + (sub) + ")";
    }

    @Override
    public Sheet clone() {
        Sheet copy = null;
        try {
            copy = (Sheet) super.clone();
        } catch (CloneNotSupportedException e) {
            ObjectOutputStream oos = null;
            ObjectInputStream ois = null;
            try {
                ByteArrayOutputStream bos = new ByteArrayOutputStream();
                oos = new ObjectOutputStream(bos);
                oos.writeObject(this);

                ois = new ObjectInputStream(new ByteArrayInputStream(bos.toByteArray()));
                copy = (Sheet) ois.readObject();
            } catch (IOException | ClassNotFoundException e1) {
                try {
                    copy = getClass().getConstructor().newInstance();
                } catch (NoSuchMethodException | IllegalAccessException
                    | InstantiationException | InvocationTargetException e2) {
                    e2.printStackTrace();
                }
            } finally {
                FileUtil.close(oos);
                FileUtil.close(ois);
            }
        }
        if (copy != null) {
            copy.copyCount = ++copyCount;
            copy.name = getCopySheetName();
            copy.relManager = relManager.deepClone();
            copy.sheetWriter = sheetWriter.clone().setWorksheet(copy);
            copy.copySheet = true;
            copy.rows = 0;
        }
        return copy;
    }

    /**
     * Check the limit of columns
     */
    public void checkColumnLimit() {
        int a = columns.length
            , b = sheetWriter.getColumnLimit();
        if (a > b) {
            throw new TooManyColumnsException(a, b);
        }
    }

    /**
     * Check the header information is exist
     *
     * @return true if exist
     */
    public boolean hasHeaderColumns() {
        return columns != null && columns.length > 0;
    }

    /**
     * Int conversion to column string number.
     * The max column on sheet is 16_384 after office 2007 and 256 in office 2003
     * <blockquote><pre>
     * int    | column number
     * -------|---------
     * 1      | A
     * 10     | J
     * 26     | Z
     * 27     | AA
     * 28     | AB
     * 53     | BA
     * 16_384 | XFD
     * </pre></blockquote>
     * @param n the column number
     * @return column string
     */
    public static char[] int2Col(int n) {
        char[][] cache_col = cache.get();
        char[] c;
        char A = 'A';
        if (n <= 26) {
            c = cache_col[0];
            c[0] = (char) (n - 1 + A);
        } else if (n <= 702) {
            int t = n / 26, w = n % 26;
            if (w == 0) {
                t--;
                w = 26;
            }
            c = cache_col[1];
            c[0] = (char) (t - 1 + A);
            c[1] = (char) (w - 1 + A);
        } else {
            int tt = n / 26, t = tt / 26, w = n % 26, m = tt % 26;
            if (w == 0) {
                m--;
                w = 26;
            }
            if (m <= 0) {
                t--;
                m += 26;
            }
            c = cache_col[2];
            c[0] = (char) (t - 1 + A);
            c[1] = (char) (m - 1 + A);
            c[2] = (char) (w - 1 + A);
        }
        return c;
    }

    private static ThreadLocal<char[][]> cache
        = ThreadLocal.withInitial(() -> new char[][]{ {65}, {65, 65}, {65, 65, 65} });

    /**
     * Check empty header row
     *
     * @return true if none header row
     */
    public boolean hasNonHeader() {
        columns = getHeaderColumns();
        boolean noneHeader = columns == null || columns.length == 0;
        if (!noneHeader) {
            int n = 0;
            for (Column column : columns) {
                if (isEmpty(column.name)) n++;
            }
            noneHeader = n == columns.length;
        }
        if (noneHeader) rows--;
        return noneHeader;
    }

    ////////////////////////////Abstract function\\\\\\\\\\\\\\\\\\\\\\\\\\\

    /**
     * Each row-block is multiplexed and will be called to reset
     * the data when a row-block is completely written.
     * Call the {@link #getRowBlockSize()} method to get
     * the row-block size, call the {@link ICellValueAndStyle#reset(int, Cell, Object, Sheet.Column)}
     * method to set value and styles.
     */
    protected abstract void resetBlockData();
}

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

package cn.ttzero.excel.entity;

import cn.ttzero.excel.annotation.TopNS;
import cn.ttzero.excel.entity.style.*;
import cn.ttzero.excel.manager.Const;
import cn.ttzero.excel.manager.RelManager;
import cn.ttzero.excel.processor.IntConversionProcessor;
import cn.ttzero.excel.processor.StyleProcessor;
import cn.ttzero.excel.reader.Cell;
import cn.ttzero.excel.util.FileUtil;

import java.io.*;
import java.lang.reflect.InvocationTargetException;
import java.math.BigDecimal;
import java.nio.file.Path;
import java.sql.Timestamp;

import static cn.ttzero.excel.entity.IWorksheetWriter.*;
import static cn.ttzero.excel.manager.Const.ROW_BLOCK_SIZE;
import static cn.ttzero.excel.util.DateUtil.toDateTimeValue;
import static cn.ttzero.excel.util.DateUtil.toDateValue;
import static cn.ttzero.excel.util.DateUtil.toTimeValue;

/**
 * Each worksheet corresponds to one or more sheet.xml of physical.
 * When the amount of data exceeds the upper limit of the worksheet,
 * the extra data will be written in the next worksheet page of the
 * current position, with the name of the parent worksheet. After
 * adding "(1,2,3...n)" as the name of the copied sheet, the pagination
 * is automatic without additional settings.
 *
 * Usually worksheetWriter calls the
 * {@code worksheet#nextBlock} method to load a row-block for writing.
 * When the row-block returns the flag EOF, mean is the current worksheet
 * finished written, and the next worksheet is written.
 *
 * Extends the existing worksheet to implement a custom data source worksheet.
 * The data source can be micro-services, Mybatis, JPA or any others. If
 * the data source returns an array of json objects, please convert to
 * an object ArrayList or Map ArrayList, the object ArrayList needs to
 * extends {@code ListSheet}, the Map ArrayList needs to extends
 * {@code ListMapSheet} and implement the {@code more} method.
 *
 * If other formats cannot be converted to ArrayList, you
 * need to inherit from the base class {@code Sheet} and implement the
 * {@code resetBlockData} and {@code getHeaderColumns} methods.
 *
 * @see ListSheet
 * @see ListMapSheet
 * @see ResultSetSheet
 * @see StatementSheet
 * Created by guanquan.wang on 2017/9/26.
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

    public int getId() {
        return id;
    }

    public void setId(int id) {
        this.id = id;
    }

    public void setSheetWriter(IWorksheetWriter sheetWriter) {
        this.sheetWriter = sheetWriter;
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
     * 伴生于Sheet用于控制头部样式和缓存列数据the cell type和转换
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
         * int值转换
         */
        public IntConversionProcessor processor;
        /**
         * 样式转换
         */
        public StyleProcessor styleProcessor;
        /**
         * 单元格样式 -1表示未设定
         */
        public int cellStyle = -1;
        /**
         * 列宽
         */
        public double width;
        public int o;
        public Styles styles;

        /**
         * 指定the column name和the cell type
         *
         * @param name  the column name
         * @param clazz the cell type
         */
        public Column(String name, Class<?> clazz) {
            this(name, clazz, true);
        }

        /**
         * 指定the column name和对应对象中的field
         *
         * @param name the column name
         * @param key  field
         */
        public Column(String name, String key) {
            this(name, key, true);
        }

        /**
         * 指定the column name对应对象中的field和the cell type，不指定类型时默认取field类型
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
         * 指定the column name，the cell type和值转换
         *
         * @param name      the column name
         * @param clazz     the cell type
         * @param processor 转换
         */
        public Column(String name, Class<?> clazz, IntConversionProcessor processor) {
            this(name, clazz, processor, true);
        }

        /**
         * 指定the column name，对象field和值转换
         *
         * @param name      the column name
         * @param key       field
         * @param processor 转换
         */
        public Column(String name, String key, IntConversionProcessor processor) {
            this(name, key, processor, true);
        }

        /**
         * 指定the column name，the cell type和是否设定值共享
         * 共享仅对字符串有效，转换后的the cell type为字符串也同样有效
         * 默认非共享以innerStr方式设值
         *
         * @param name  the column name
         * @param clazz the cell type
         * @param share true:共享 false:非共享
         */
        public Column(String name, Class<?> clazz, boolean share) {
            this.name = name;
            this.clazz = clazz;
            this.share = share;
        }

        /**
         * 指定the column name，field和是否共享字串
         * 共享仅对字符串有效，转换后的the cell type为字符串也同样有效
         * 默认非共享以innerStr方式设值
         *
         * @param name  the column name
         * @param key   filed
         * @param share true:共享 false:非共享
         */
        public Column(String name, String key, boolean share) {
            this.name = name;
            this.key = key;
            this.share = share;
        }

        /**
         * 指定the column name，the cell type，转换和是否共享字串
         * 共享仅对字符串有效，转换后的the cell type为字符串也同样有效
         * 默认非共享以innerStr方式设值
         *
         * @param name      the column name
         * @param clazz     the cell type
         * @param processor 转换
         * @param share     true:共享 false:非共享
         */
        public Column(String name, Class<?> clazz, IntConversionProcessor processor, boolean share) {
            this(name, clazz, share);
            this.processor = processor;
        }

        /**
         * 指定the column name，field，转换和是否共享字串
         * 共享仅对字符串有效，转换后的the cell type为字符串也同样有效
         * 默认非共享以innerStr方式设值
         *
         * @param name      the column name
         * @param key       field
         * @param clazz     type of cell
         * @param processor 转换
         */
        public Column(String name, String key, Class<?> clazz, IntConversionProcessor processor) {
            this(name, key, clazz);
            this.processor = processor;
        }

        /**
         * 指定the column name，field，转换和是否共享字串
         * 共享仅对字符串有效，转换后的the cell type为字符串也同样有效
         * 默认非共享以innerStr方式设值
         *
         * @param name      the column name
         * @param key       field
         * @param processor 转换
         * @param share     true:共享 false:非共享
         */
        public Column(String name, String key, IntConversionProcessor processor, boolean share) {
            this(name, key, share);
            this.processor = processor;
        }

        /**
         * 指定the column name，the cell type和单元样式
         *
         * @param name      the column name
         * @param clazz     the cell type
         * @param cellStyle 样式
         */
        public Column(String name, Class<?> clazz, int cellStyle) {
            this(name, clazz, cellStyle, true);
        }

        /**
         * 指定列，field 和单元样式
         *
         * @param name      the column name
         * @param key       field
         * @param cellStyle 样式
         */
        public Column(String name, String key, int cellStyle) {
            this(name, key, cellStyle, true);
        }

        /**
         * 指事实上the column name，the cell type，样式以及是否共享
         * 共享仅对字符串有效，转换后的the cell type为字符串也同样有效
         * 默认非共享以innerStr方式设值
         *
         * @param name      the column name
         * @param clazz     the cell type
         * @param cellStyle 样式
         * @param share     true:共享 false:非共享
         */
        public Column(String name, Class<?> clazz, int cellStyle, boolean share) {
            this(name, clazz, share);
            this.cellStyle = cellStyle;
        }

        /**
         * 指事实上the column name，field，样式以及是否共享
         * 共享仅对字符串有效，转换后的the cell type为字符串也同样有效
         * 默认非共享以innerStr方式设值
         *
         * @param name      the column name
         * @param key       field
         * @param cellStyle 样式
         * @param share     true:共享 false:非共享
         */
        public Column(String name, String key, int cellStyle, boolean share) {
            this(name, key, share);
            this.cellStyle = cellStyle;
        }

        /**
         * 设定单元格宽度
         */
        public Column setWidth(double width) {
            if (width < 0.00000001) {
                throw new ExcelWriteException("Width " + width + " less than 0.");
            }
            this.width = width;
            return this;
        }

        /**
         * 单元格共享
         *
         * @return true:共享 false:非共享
         */
        public boolean isShare() {
            return share;
        }

        /**
         * 设置单元格the cell type
         *
         * @param type the cell type
         * @return Column实例
         * @see Const.ColumnType
         */
        public Column setType(int type) {
            this.type = type;
            return this;
        }

        /**
         * 获取列头名
         *
         * @return the column name
         */
        public String getName() {
            return name;
        }

        /**
         * 设置the column name
         *
         * @param name the column name
         * @return Column实例
         */
        public Column setName(String name) {
            this.name = name;
            return this;
        }

        /**
         * 获取列the cell type
         *
         * @return the cell type
         */
        public Class<?> getClazz() {
            return clazz;
        }

        /**
         * 设置列the cell type
         *
         * @param clazz the cell type
         * @return Column实例
         */
        public Column setClazz(Class<?> clazz) {
            this.clazz = clazz;
            return this;
        }

        /**
         * 设置值转换
         * 每个单元格只能有一个值转换，多次set后最后一个有效
         *
         * @param processor 值转换
         * @return Column实例
         */
        public Column setProcessor(IntConversionProcessor processor) {
            this.processor = processor;
            return this;
        }

        /**
         * 设置样式转换
         * 每个单元格只能有一个样式转换，多次set后最后一个有效
         *
         * @param styleProcessor 样式转换
         * @return Column实例
         */
        public Column setStyleProcessor(StyleProcessor styleProcessor) {
            this.styleProcessor = styleProcessor;
            return this;
        }

        /**
         * 获得列宽
         *
         * @return 列宽
         */
        public double getWidth() {
            return width;
        }

        /**
         * 设置样式
         * 样式值必须是调用style.add获取的值
         *
         * @param cellStyle 样式值
         * @return Column实例
         */
        public Column setCellStyle(int cellStyle) {
            this.cellStyle = cellStyle;
            return this;
        }

        /**
         * 默认水平对齐
         * 日期，字符，bool值居中，numeric居右其余居左
         *
         * @return int
         * @see Horizontals
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
         * 设置单元格样式
         *
         * @param font 字体
         * @return Column实例
         */
        public Column setCellStyle(Font font) {
            this.cellStyle = styles.of(
                (font != null ? styles.addFont(font) : 0)
                    | Verticals.CENTER
                    | defaultHorizontal());
            return this;
        }

        /**
         * 设置单元格样式
         *
         * @param font       字体
         * @param horizontal 水平对齐
         * @return Column实例
         */
        public Column setCellStyle(Font font, int horizontal) {
            this.cellStyle = styles.of(
                (font != null ? styles.addFont(font) : 0)
                    | Verticals.CENTER
                    | horizontal);
            return this;
        }

        /**
         * 设置单元格样式
         *
         * @param font   字体
         * @param border 边框
         * @return Column实例
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
         * 设置单元格样式
         *
         * @param font       字体
         * @param border     边框
         * @param horizontal 水平对齐
         * @return
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
         * 设置单元格样式
         *
         * @param font   字体
         * @param fill   填充
         * @param border 边框
         * @return Column实例
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
         * 设置单元格样式
         *
         * @param font       字体
         * @param fill       填充
         * @param border     边框
         * @param horizontal 水平对齐
         * @return Column实例
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
         * 设置单元格样式
         *
         * @param font       字体
         * @param fill       填充
         * @param border     边框
         * @param vertical   垂直对齐
         * @param horizontal 水平对齐
         * @return Column实例
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
         * 设置单元格样式
         *
         * @param numFmt     格式化
         * @param font       字体
         * @param fill       填充
         * @param border     边框
         * @param vertical   垂直对齐
         * @param horizontal 水平对齐
         * @return Column实例
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
         * 设置共享
         * 共享仅对字符串有效，转换后的the cell type为字符串也同样有效
         * 默认非共享以innerStr方式设值
         *
         * @param share true:共享 false:非共享
         * @return Column实例
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
         * 根据the cell type获取默认样式
         *
         * @param clazz the cell type
         * @return 样式
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
         * 获取单元格样式
         *
         * @return
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
     * @return the workbook
     */
    public Workbook getWorkbook() {
        return workbook;
    }

    /**
     * Setting worksheet
     *
     * @param workbook the worksheet
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
     * Output export info
     */
    public void what(String code) {
        workbook.what(code);
    }

    /**
     * Output export info
     */
    public void what(String code, String... args) {
        workbook.what(code, args);
    }

    /**
     * Returns shared string
     *
     * @return the sst
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
     * 设置列宽自动调节
     *
     * @return worksheet实例
     */
    public Sheet autoSize() {
        this.autoSize = 1;
        return this;
    }

    /**
     * 固定列宽，默认20
     *
     * @return worksheet实例
     */
    public Sheet fixSize() {
        this.autoSize = 2;
        return this;
    }

    /**
     * 指定固定列宽
     *
     * @param width 列宽
     * @return worksheet实例
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
     * 获取是否自动调节列宽
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
     * 取消隔行变色
     *
     * @return worksheet实例
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
     * 设置偶数行填充颜色
     *
     * @param fill 填充
     * @return worksheet实例
     */
    public Sheet setOddFill(Fill fill) {
        this.oddFill = workbook.getStyles().addFill(fill);
        return this;
    }

    /**
     * 获取worksheet名
     *
     * @return worksheet名
     */
    public String getName() {
        return name;
    }

    /**
     * 设置worksheet名
     *
     * @param name sheet名
     * @return worksheet实例
     */
    public Sheet setName(String name) {
        this.name = name;
        return this;
    }

    /**
     * Returns the header column info
     *
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
     * 设置列头
     *
     * @param columns 列头数组
     * @return worksheet实例
     */
    public Sheet setColumns(final Column[] columns) {
        this.columns = columns.clone();
        for (int i = 0; i < columns.length; i++) {
            columns[i].styles = workbook.getStyles();
        }
        return this;
    }

    /**
     * 获取水印
     *
     * @return 水印
     * @see WaterMark
     */
    public WaterMark getWaterMark() {
        return waterMark;
    }

    /**
     * 设置水印
     *
     * @param waterMark 水印
     * @return worksheet实例
     */
    public Sheet setWaterMark(WaterMark waterMark) {
        this.waterMark = waterMark;
        return this;
    }

    /**
     * 单元是否隐藏
     *
     * @return true: hidden, false: not hidden
     */
    public boolean isHidden() {
        return hidden;
    }

    /**
     * 设置单元格隐藏
     *
     * @return worksheet实例
     */
    public Sheet hidden() {
        this.hidden = true;
        return this;
    }

    /**
     * abstract method close
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
     * @throws IOException         write error
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

    public String getFileName() {
        return "sheet" + id + Const.Suffix.XML;
    }

    /**
     * 设置表头样式
     *
     * @param font   字体
     * @param fill   填充
     * @param border 边框
     * @return worksheet实例
     */
    public Sheet setHeadStyle(Font font, Fill fill, Border border) {
        return setHeadStyle(null, font, fill, border, Verticals.CENTER, Horizontals.CENTER);
    }

    /**
     * 设置表头样式
     *
     * @param font       字体
     * @param fill       填允
     * @param border     边框
     * @param vertical   垂直对齐
     * @param horizontal 水平对齐
     * @return worksheet实例
     */
    public Sheet setHeadStyle(Font font, Fill fill, Border border, int vertical, int horizontal) {
        return setHeadStyle(null, font, fill, border, vertical, horizontal);
    }

    /**
     * 设置表头样式
     *
     * @param numFmt     格式化
     * @param font       字体
     * @param fill       填充
     * @param border     边框
     * @param vertical   垂直对齐
     * @param horizontal 水平对齐
     * @return worksheet实例
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
     * 样式表头样式
     *
     * @param style 样式值
     * @return worksheet实例
     */
    public Sheet setHeadStyle(int style) {
        headStyle = style;
        return this;
    }

    public int defaultHeadStyle() {
        if (headStyle == 0) {
            Styles styles = workbook.getStyles();
            headStyle = styles.of(styles.addFont(Font.parse("bold 11 微软雅黑 white"))
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
     * written at a time. If the data is not enough, the {@code more}
     * method will be called to get more data.
     *
     * @return the row-block size
     */
    public int getRowBlockSize() {
        return ROW_BLOCK_SIZE;
    }

    /**
     * Write some final info
     */
    public void afterSheetAccess(Path workSheetPath) throws IOException {
        // relationship
        relManager.write(workSheetPath, getFileName());

        // others ...
    }

    /**
     * Int value conversion to others
     *
     * @param cell the cell
     * @param n    the cell value
     * @param hc   the header column
     */
    protected void conversion(Cell cell, int n, Column hc) {
        Object e = hc.processor.conversion(n);
        if (e != null) {
            Class<?> clazz = e.getClass();
            if (isInt(clazz)) {
                if (isChar(clazz)) {
                    cell.setCv((Character) e);
                } else if (isShort(clazz)) {
                    cell.setNv((Short) e);
                } else {
                    cell.setNv((Integer) e);
                }
                cell.xf = getStyleIndex(hc, e);
            } else {
                setCellValue(cell, e, hc, clazz);
                int style = hc.getCellStyle(clazz);
                cell.xf = getStyleIndex(hc, n, style);
            }
        } else {
            cell.setBlank();
            cell.xf = getStyleIndex(hc, e);
        }
    }

    /**
     * Setting cell value and cell styles
     *
     * @param cell the cell
     * @param e    the cell value
     * @param hc   the header column
     */
    protected void setCellValueAndStyle(Cell cell, Object e, Column hc) {
        setCellValue(cell, e, hc, hc.getClazz());
        if (hc.processor == null) {
            cell.xf = getStyleIndex(hc, e);
        }
    }

    /**
     * Setting cell value
     *
     * @param cell  the cell
     * @param e     the cell value
     * @param hc    the header column
     * @param clazz the cell value type
     */
    protected void setCellValue(Cell cell, Object e, Column hc, Class<?> clazz) {
        boolean hasIntProcessor = hc.processor != null;
        if (isString(clazz)) {
            cell.setSv(e.toString());
        } else if (isDate(clazz)) {
            cell.setAv(toDateValue((java.util.Date) e));
        } else if (isDateTime(clazz)) {
            cell.setIv(toDateTimeValue((Timestamp) e));
        } else if (isChar(clazz)) {
            Character c = (Character) e;
            if (hasIntProcessor) conversion(cell, c, hc);
            else cell.setCv(c);
        } else if (isShort(clazz)) {
            Short t = (Short) e;
            if (hasIntProcessor) conversion(cell, t, hc);
            else cell.setNv(t);
        } else if (isInt(clazz)) {
            Integer n = (Integer) e;
            if (hasIntProcessor) conversion(cell, n, hc);
            else cell.setNv(n);
        } else if (isLong(clazz)) {
            cell.setLv((Long) e);
        } else if (isFloat(clazz)) {
            cell.setDv((Float) e);
        } else if (isDouble(clazz)) {
            cell.setDv((Double) e);
        } else if (isBool(clazz)) {
            cell.setBv((Boolean) e);
        } else if (isBigDecimal(clazz)) {
            cell.setMv((BigDecimal) e);
        } else if (isLocalDate(clazz)) {
            cell.setAv(toDateValue((java.time.LocalDate) e));
        } else if (isLocalDateTime(clazz)) {
            cell.setIv(toDateTimeValue((java.time.LocalDateTime) e));
        } else if (isTime(clazz)) {
            cell.setTv(toTimeValue((java.sql.Time) e));
        } else if (isLocalTime(clazz)) {
            cell.setTv(toTimeValue((java.time.LocalTime) e));
        } else {
            cell.setSv(e.toString());
        }
    }

    /**
     * Returns the cell style index
     *
     * @param hc    the header column
     * @param o     the cell value
     * @param style the default style
     * @return the style index in xf
     */
    private int getStyleIndex(Column hc, Object o, int style) {
        // Interlaced discoloration
        if (autoOdd == 0 && isOdd() && !Styles.hasFill(style)) {
            style |= oddFill;
        }
        int styleIndex = hc.styles.of(style);
        if (hc.styleProcessor != null) {
            style = hc.styleProcessor.build(o, style, hc.styles);
            styleIndex = hc.styles.of(style);
        }
        return styleIndex;
    }

    /**
     * Returns the cell style index
     *
     * @param hc the header column
     * @param o  the cell value
     * @return the style index in xf
     */
    protected int getStyleIndex(Column hc, Object o) {
        int style = hc.getCellStyle();
        return getStyleIndex(hc, o, style);
    }

    /**
     * Check the odd rows
     * @return true if odd rows
     */
    protected boolean isOdd() {
        return (rows & 1) == 1;
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
            copy.relManager = relManager.clone();
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
     * @return true if exist
     */
    public boolean hasHeaderColumns() {
        return columns != null && columns.length > 0;
    }

    /**
     * Int conversion to column string number
     * The max column on sheet is 16_384
     * <p>
     * int  | column number
     * -------|---------
     * 1      | A
     * 10     | J
     * 26     | Z
     * 27     | AA
     * 28     | AB
     * 53     | BA
     * 16_384 | XFD
     */
    private ThreadLocal<char[][]> cache = ThreadLocal.withInitial(() -> new char[][]{{65}, {65, 65}, {65, 65, 65}});

    public char[] int2Col(int n) {
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

    ////////////////////////////Abstract function\\\\\\\\\\\\\\\\\\\\\\\\\\\

    /**
     * Each row-block is multiplexed and will be called to reset
     * the data when a row-block is completely written.
     * Call the {@code RowBlock#getRowBlockSize} method to get
     * the row-block size, call the {@code setCellValueAndStyle}
     * method to set value and styles.
     */
    protected abstract void resetBlockData();
}

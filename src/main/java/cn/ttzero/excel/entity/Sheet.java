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

import java.io.*;
import java.math.BigDecimal;
import java.nio.file.Path;
import java.sql.Timestamp;

import static cn.ttzero.excel.entity.IWorksheetWriter.*;
import static cn.ttzero.excel.util.DateUtil.toDateTimeValue;
import static cn.ttzero.excel.util.DateUtil.toDateValue;
import static cn.ttzero.excel.util.DateUtil.toTimeValue;

/**
 * 对应workbook各sheet页
 * Created by guanquan.wang on 2017/9/26.
 */
@TopNS(prefix = {"", "r"}, value = "worksheet", uri = {Const.SCHEMA_MAIN, Const.Relationship.RELATIONSHIP})
public abstract class Sheet {
    protected Workbook workbook;

    protected String name;
    protected Column[] columns;
    protected WaterMark waterMark;
    protected RelManager relManager;
    protected int id;
    /** 自动列宽 */
    protected int autoSize;
    /** 默认固定宽度 */
    protected double width = 20;
    /** 记录已写入行数，每写一行该数加1 */
    protected int rows;
    /** 设置单元格默认隐藏 */
    protected boolean hidden;

    protected int headStyle;
    /** 自动隔行变色 */
    protected int autoOdd = -1;
    /** odd row's background color */
    protected int oddFill;
    protected boolean copySheet;

    protected RowBlock rowBlock;
    protected IWorksheetWriter sheetWriter;
    /** To mark the header column is ready */
    protected boolean headerReady;
    /** Close resource on the last copy worksheet */
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
     * @param name the worksheet name
     */
    public Sheet(String name) {
        this.name = name;
        relManager = new RelManager();
    }

    /**
     * Constructor worksheet
     * @param name the worksheet name
     * @param columns the header info
     */
    public Sheet(String name, final Column[] columns) {
        this(name, null, columns);
    }

    /**
     * Constructor worksheet
     * @param name the worksheet name
     * @param waterMark the water mark
     * @param columns the header info
     */
    public Sheet(String name, WaterMark waterMark, final Column[] columns) {
        this.name = name;
        this.columns = columns;
        this.waterMark = waterMark;
        relManager = new RelManager();
    }

    /**
     * 伴生于Sheet用于控制头部样式和缓存列数据类型和转换
     */
    public static class Column {
        /** Map的主键,object的属性名 */
        public String key;
        /** 列头名 */
        public String name;
        /** 列类型 */
        public Class<?> clazz;
        /** 字符串是否共享 */
        public boolean share;
        /** 0: 正常显示 1:显示百分比 2:显示人民币 */
        public int type;
        /** int值转换 */
        public IntConversionProcessor processor;
        /** 样式转换 */
        public StyleProcessor styleProcessor;
        /** 单元格样式 -1表示未设定 */
        public int cellStyle = -1;
        /** 列宽 */
        public double width;
        public Object o;
        public Styles styles;

        /**
         * 指定列名和类型
         * @param name 列名
         * @param clazz 类型
         */
        public Column(String name, Class<?> clazz) {
            this(name, clazz, false);
        }

        /**
         * 指定列名和对应对象中的field
         * @param name 列名
         * @param key field
         */
        public Column(String name, String key) {
            this(name, key, false);
        }

        /**
         * 指定列名对应对象中的field和类型，不指定类型时默认取field类型
         * @param name 列名
         * @param key field
         * @param clazz 类型
         */
        public Column(String name, String key, Class<?> clazz) {
            this(name, key, false);
            this.clazz = clazz;
        }

        /**
         * 指定列名，类型和值转换
         * @param name 列名
         * @param clazz 类型
         * @param processor 转换
         */
        public Column(String name, Class<?> clazz, IntConversionProcessor processor) {
            this(name, clazz, processor, false);
        }

        /**
         * 指定列名，对象field和值转换
         * @param name 列名
         * @param key field
         * @param processor 转换
         */
        public Column(String name, String key, IntConversionProcessor processor) {
            this(name, key, processor, false);
        }

        /**
         * 指定列名，类型和是否设定值共享
         * 共享仅对字符串有效，转换后的类型为字符串也同样有效
         * 默认非共享以innerStr方式设值
         * @param name 列名
         * @param clazz 类型
         * @param share true:共享 false:非共享
         */
        public Column(String name, Class<?> clazz, boolean share) {
            this.name = name;
            this.clazz = clazz;
            this.share = share;
        }

        /**
         * 指定列名，field和是否共享字串
         * 共享仅对字符串有效，转换后的类型为字符串也同样有效
         * 默认非共享以innerStr方式设值
         * @param name 列名
         * @param key filed
         * @param share true:共享 false:非共享
         */
        public Column(String name, String key, boolean share) {
            this.name = name;
            this.key = key;
            this.share = share;
        }

        /**
         * 指定列名，类型，转换和是否共享字串
         * 共享仅对字符串有效，转换后的类型为字符串也同样有效
         * 默认非共享以innerStr方式设值
         * @param name 列名
         * @param clazz 类型
         * @param processor 转换
         * @param share true:共享 false:非共享
         */
        public Column(String name, Class<?> clazz, IntConversionProcessor processor, boolean share) {
            this(name, clazz, share);
            this.processor = processor;
        }

        /**
         * 指定列名，field，转换和是否共享字串
         * 共享仅对字符串有效，转换后的类型为字符串也同样有效
         * 默认非共享以innerStr方式设值
         * @param name 列名
         * @param key field
         * @param processor 转换
         * @param share true:共享 false:非共享
         */
        public Column(String name, String key, IntConversionProcessor processor, boolean share) {
            this(name, key, share);
            this.processor = processor;
        }

        /**
         * 指定列名，类型和单元样式
         * @param name 列名
         * @param clazz 类型
         * @param cellStyle 样式
         */
        public Column(String name, Class<?> clazz, int cellStyle) {
            this(name, clazz, cellStyle, false);
        }
        /**
         * 指定列，field 和单元样式
         * @param name 列名
         * @param key field
         * @param cellStyle 样式
         */
        public Column(String name, String key, int cellStyle) {
            this(name, key, cellStyle, false);
        }

        /**
         * 指事实上列名，类型，样式以及是否共享
         * 共享仅对字符串有效，转换后的类型为字符串也同样有效
         * 默认非共享以innerStr方式设值
         * @param name 列名
         * @param clazz 类型
         * @param cellStyle 样式
         * @param share true:共享 false:非共享
         */
        public Column(String name, Class<?> clazz, int cellStyle, boolean share) {
            this(name, clazz, share);
            this.cellStyle = cellStyle;
        }
        /**
         * 指事实上列名，field，样式以及是否共享
         * 共享仅对字符串有效，转换后的类型为字符串也同样有效
         * 默认非共享以innerStr方式设值
         * @param name 列名
         * @param key field
         * @param cellStyle 样式
         * @param share true:共享 false:非共享
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
         * @return true:共享 false:非共享
         */
        public boolean isShare() {
            return share;
        }

        /**
         * 设置单元格类型
         * @see Const.ColumnType
         * @param type 类型
         * @return Column实例
         */
        public Column setType(int type) {
            this.type = type;
            return this;
        }

        /**
         * 获取列头名
         * @return 列名
         */
        public String getName() {
            return name;
        }

        /**
         * 设置列名
         * @param name 列名
         * @return Column实例
         */
        public Column setName(String name) {
            this.name = name;
            return this;
        }

        /**
         * 获取列类型
         * @return 类型
         */
        public Class<?> getClazz() {
            return clazz;
        }

        /**
         * 设置列类型
         * @param clazz 类型
         * @return Column实例
         */
        public Column setClazz(Class<?> clazz) {
            this.clazz = clazz;
            return this;
        }

        /**
         * 设置值转换
         * 每个单元格只能有一个值转换，多次set后最后一个有效
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
         * @param styleProcessor 样式转换
         * @return Column实例
         */
        public Column setStyleProcessor(StyleProcessor styleProcessor) {
            this.styleProcessor = styleProcessor;
            return this;
        }

        /**
         * 获得列宽
         * @return 列宽
         */
        public double getWidth() {
            return width;
        }

        /**
         * 设置样式
         * 样式值必须是调用style.add获取的值
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
         * @see Horizontals
         * @return int
         */
        int defaultHorizontal() {
            int horizontal;
            if (isDate(clazz) || isDateTime(clazz)
                    || isLocalDate(clazz) || isLocalDateTime(clazz)
                    || isTime(clazz) || isLocalTime(clazz)
                    || isChar(clazz) || isBool(clazz)) {
                horizontal = Horizontals.CENTER;
            } else if (isInt(clazz) || isLong(clazz) || isFloat(clazz) || isBigDecimal(clazz)) {
                horizontal = Horizontals.RIGHT;
            } else {
                horizontal = Horizontals.LEFT;
            }
            return horizontal;
        }

        /**
         * 设置单元格样式
         * @param font 字体
         * @return Column实例
         */
        public Column setCellStyle(Font font) {
            this.cellStyle =  styles.of(
                    (font != null ? styles.addFont(font) : 0)
                            | Verticals.CENTER
                            | defaultHorizontal());
            return this;
        }

        /**
         * 设置单元格样式
         * @param font 字体
         * @param horizontal 水平对齐
         * @return Column实例
         */
        public Column setCellStyle(Font font, int horizontal) {
            this.cellStyle =  styles.of(
                    (font != null ? styles.addFont(font) : 0)
                            | Verticals.CENTER
                            | horizontal);
            return this;
        }

        /**
         * 设置单元格样式
         * @param font 字体
         * @param border 边框
         * @return Column实例
         */
        public Column setCellStyle(Font font, Border border) {
            this.cellStyle =  styles.of(
                    (font != null ? styles.addFont(font) : 0)
                            | (border != null ? styles.addBorder(border) : 0)
                            | Verticals.CENTER
                            | defaultHorizontal());
            return this;
        }

        /**
         * 设置单元格样式
         * @param font 字体
         * @param border 边框
         * @param horizontal 水平对齐
         * @return
         */
        public Column setCellStyle(Font font, Border border, int horizontal) {
            this.cellStyle =  styles.of(
                    (font != null ? styles.addFont(font) : 0)
                            | (border != null ? styles.addBorder(border) : 0)
                            | Verticals.CENTER
                            | horizontal);
            return this;
        }

        /**
         * 设置单元格样式
         * @param font 字体
         * @param fill 填充
         * @param border 边框
         * @return Column实例
         */
        public Column setCellStyle(Font font, Fill fill, Border border) {
            this.cellStyle =  styles.of(
                    (font != null ? styles.addFont(font) : 0)
                            | (fill != null ? styles.addFill(fill) : 0)
                            | (border != null ? styles.addBorder(border) : 0)
                            | Verticals.CENTER
                            | defaultHorizontal());
            return this;
        }

        /**
         * 设置单元格样式
         * @param font 字体
         * @param fill 填充
         * @param border 边框
         * @param horizontal 水平对齐
         * @return Column实例
         */
        public Column setCellStyle(Font font, Fill fill, Border border, int horizontal) {
            this.cellStyle =  styles.of(
                    (font != null ? styles.addFont(font) : 0)
                            | (fill != null ? styles.addFill(fill) : 0)
                            | (border != null ? styles.addBorder(border) : 0)
                            | Verticals.CENTER
                            | horizontal);
            return this;
        }

        /**
         * 设置单元格样式
         * @param font 字体
         * @param fill 填充
         * @param border 边框
         * @param vertical 垂直对齐
         * @param horizontal 水平对齐
         * @return Column实例
         */
        public Column setCellStyle(Font font, Fill fill, Border border, int vertical, int horizontal) {
            this.cellStyle =  styles.of(
                            (font != null ? styles.addFont(font) : 0)
                            | (fill != null ? styles.addFill(fill) : 0)
                            | (border != null ? styles.addBorder(border) : 0)
                            | vertical
                            | horizontal);
            return this;
        }

        /**
         * 设置单元格样式
         * @param numFmt 格式化
         * @param font 字体
         * @param fill 填充
         * @param border 边框
         * @param vertical 垂直对齐
         * @param horizontal 水平对齐
         * @return Column实例
         */
        public Column setCellStyle(NumFmt numFmt, Font font, Fill fill, Border border, int vertical, int horizontal) {
            this.cellStyle =  styles.of(
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
         * 共享仅对字符串有效，转换后的类型为字符串也同样有效
         * 默认非共享以innerStr方式设值
         * @param share true:共享 false:非共享
         * @return Column实例
         */
        public Column setShare(boolean share) {
            this.share = share;
            return this;
        }

        private NumFmt ip = new NumFmt("0%_);[Red]\\(0%\\)") // 整数百分比
                , ir = new NumFmt("¥0_);[Red]\\(¥0\\)") // 整数人民币
                , fp = new NumFmt("0.00%_);[Red]\\(0.00%\\)") // 小数百分比
                , fr = new NumFmt("¥0.00_);[Red]\\(¥0.00\\)") // 小数人民币
                , tm = new NumFmt("hh:mm:ss") // 时分秒
                ;

        /**
         * 根据类型获取默认样式
         * @param clazz 数据类型
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
                        style= Styles.clearNumfmt(style) | styles.addNumFmt(fp);
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
     * @return the workbook
     */
    public Workbook getWorkbook() {
        return workbook;
    }

    /**
     * Setting worksheet
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
     * Return the cell default width
     * @return the width value
     */
    public double getDefaultWidth() {
        return width;
    }

    /**
     * 设置列宽自动调节
     * @return worksheet实例
     */
    public Sheet autoSize() {
        this.autoSize = 1;
        return this;
    }

    /**
     * 固定列宽，默认20
     * @return worksheet实例
     */
    public Sheet fixSize() {
        this.autoSize = 2;
        return this;
    }

    /**
     * 指定固定列宽
     * @param width 列宽
     * @return worksheet实例
     */
    public Sheet fixSize(double width) {
        this.autoSize = 2;
        for (Column hc : columns) {
            hc.setWidth(width);
        }
        return this;
    }

    /**
     * 获取是否自动调节列宽
     * @return 1: auto-size 2:fix-size
     */
    public int getAutoSize() {
        return autoSize;
    }

    /**
     * Test is auto size column width
     * @return true if auto-size
     */
    public boolean isAutoSize() {
        return autoSize == 1;
    }

    /**
     * 取消隔行变色
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
     * @param fill 填充
     * @return worksheet实例
     */
    public Sheet setOddFill(Fill fill) {
        this.oddFill = workbook.getStyles().addFill(fill);
        return this;
    }

    /**
     * 获取worksheet名
     * @return worksheet名
     */
    public String getName() {
        return name;
    }

    /**
     * 设置worksheet名
     * @param name sheet名
     * @return worksheet实例
     */
    public Sheet setName(String name) {
        this.name = name;
        return this;
    }

    /**
     * Returns the header column info
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
     * @see WaterMark
     * @return 水印
     */
    public WaterMark getWaterMark() {
        return waterMark;
    }

    /**
     * 设置水印
     * @param waterMark 水印
     * @return worksheet实例
     */
    public Sheet setWaterMark(WaterMark waterMark) {
        this.waterMark = waterMark;
        return this;
    }

    /**
     * 单元是否隐藏
     * @return true: hidden, false: not hidden
     */
    public boolean isHidden() {
        return hidden;
    }

    /**
     * 设置单元格隐藏
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
     * write worksheet data to path
     * @param path the storage path
     * @throws IOException write error
     * @throws ExcelWriteException others
     */
    public void writeTo(Path path) throws IOException, ExcelWriteException {
        if (sheetWriter == null) {
            throw new ExcelWriteException("Worksheet writer is not instanced.");
        }
        if (!copySheet) {
            paging();
        }
        rowBlock = new RowBlock();
        sheetWriter.write(path);
    }

    /**
     * Split worksheet data
     */
    protected void paging() { }

    /**
     * 添加关联
     * @param rel Relationship
     * @return worksheet实例
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
     * @param font 字体
     * @param fill 填充
     * @param border 边框
     * @return worksheet实例
     */
    public Sheet setHeadStyle(Font font, Fill fill, Border border) {
        return setHeadStyle(null, font, fill, border, Verticals.CENTER, Horizontals.CENTER);
    }

    /**
     * 设置表头样式
     * @param font 字体
     * @param fill 填允
     * @param border 边框
     * @param vertical 垂直对齐
     * @param horizontal 水平对齐
     * @return worksheet实例
     */
    public Sheet setHeadStyle(Font font, Fill fill, Border border, int vertical, int horizontal) {
        return setHeadStyle(null, font, fill, border, vertical, horizontal);
    }

    /**
     * 设置表头样式
     * @param numFmt 格式化
     * @param font 字体
     * @param fill 填充
     * @param border 边框
     * @param vertical 垂直对齐
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

    protected static boolean blockOrDefault(int style) {
        return style == -1 || style == Styles.defaultIntBorderStyle();
    }

    /**
     * Returns total rows in this worksheet
     * @return -1 if unknown
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
     * Write some final info
     */
    public void afterSheetAccess(Path workSheetPath) throws IOException {
        // relationship
        relManager.write(workSheetPath, getFileName());
    }


    protected void convert(Cell cell, Column hc, int n) {
        Object e = hc.processor.conversion(n);
        if (e != null) {
            setCellValue(cell, e, hc);
        } else cell.setBlank();
    }

    protected void setCellValue(Cell cell, Object e, Column hc) {
        Class<?> clazz = hc.getClazz();
        if (isString(clazz)) {
            cell.setSv(e.toString());
        } else if (isDate(clazz)) {
            cell.setAv(toDateValue((java.util.Date) e));
        } else if (isDateTime(clazz)) {
            cell.setIv(toDateTimeValue((Timestamp) e));
        } else if (isChar(clazz)) {
            cell.setCv((Character) e);
        } else if (isShort(clazz)) {
            cell.setNv((Short) e);
        } else if (isInt(clazz)) {
            cell.setNv((Integer) e);
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
        cell.xf = getStyleIndex(hc, e);
    }

    protected int getStyleIndex(Sheet.Column hc, Object o) {
        int style = hc.getCellStyle();
        // 隔行变色
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

    protected boolean isOdd() {
        return (rows & 1) == 1;
    }

    /**
     * Int转列号A-Z
     */
    private ThreadLocal<char[][]> cache = ThreadLocal.withInitial(() -> new char[][] {{65}, {65, 65}, {65, 65, 65}});
    public char[] int2Col(int n) {
        char[][] cache_col = cache.get();
        char[] c; char A = 'A';
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
     * Reset the row-block data
     */
    protected abstract void resetBlockData();
}

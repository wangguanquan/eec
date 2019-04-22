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

package cn.ttzero.excel.entity.e7;

import cn.ttzero.excel.entity.style.*;
import cn.ttzero.excel.manager.Const;
import cn.ttzero.excel.manager.RelManager;
import cn.ttzero.excel.annotation.TopNS;
import cn.ttzero.excel.entity.ExportException;
import cn.ttzero.excel.entity.WaterMark;
import cn.ttzero.excel.processor.IntConversionProcessor;
import cn.ttzero.excel.processor.StyleProcessor;
import cn.ttzero.excel.util.ExtBufferedWriter;
import cn.ttzero.excel.util.StringUtil;

import java.io.*;
import java.math.BigDecimal;
import java.nio.ByteBuffer;
import java.nio.channels.FileChannel;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Time;
import java.sql.Timestamp;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.util.Date;

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
    private int autoSize;
    /** 默认固定宽度 */
    private double width = 20;
    private int headInfoLen, baseInfoLen;
    /** 记录已写入行数，每写一行该数加1 */
    protected int rows;
    /** 设置单元格默认隐藏 */
    private boolean hidden;

    private int headStyle;
    /** 自动隔行变色 */
    protected int autoOdd = -1;
    /** odd row's background color */
    protected int oddFill;

    public int getId() {
        return id;
    }

    void setId(int id) {
        this.id = id;
    }

    /**
     * 实例化Sheet
     * @param workbook workbook
     */
    public Sheet(Workbook workbook) {
        this.workbook = workbook;
        relManager = new RelManager();
    }
    /**
     * 实例化Sheet
     * @param workbook workbook
     * @param name sheet名
     * @param columns 行信息
     */
    public Sheet(Workbook workbook, String name, final Column[] columns) {
        this.workbook = workbook;
        this.name = name;
        this.columns = columns;
        for (int i = 0; i < columns.length; i++) {
            columns[i].styles = workbook.getStyles();
        }
        relManager = new RelManager();
    }

    /**
     * 实例化Sheet
     * @param workbook workbook
     * @param name sheet名
     * @param waterMark 水印
     * @param columns 行信息
     */
    public Sheet(Workbook workbook, String name, WaterMark waterMark, final Column[] columns) {
        this.workbook = workbook;
        this.name = name;
        this.columns = columns;
        for (int i = 0; i < columns.length; i++) {
            columns[i].styles = workbook.getStyles();
        }
        this.waterMark = waterMark;
        relManager = new RelManager();
    }
    /**
     * 伴生于Sheet用于控制头部样式和缓存列数据类型和转换
     */
    public static class Column {
        /** Map的主键,object的属性名 */
        private String key;
        /** 列头名 */
        private String name;
        /** 列类型 */
        private Class<?> clazz;
        /** 字符串是否共享 */
        private boolean share;
        /** 0: 正常显示 1:显示百分比 2:显示人民币 */
        private int type;
        /** int值转换 */
        private IntConversionProcessor processor;
        /** 样式转换 */
        private StyleProcessor styleProcessor;
        /** 单元格样式 -1表示未设定 */
        private int cellStyle = -1;
        /** 列宽 */
        private double width;
        private Object o;
        private Styles styles;

        private Column setO(Object o) {
            this.o = o;
            return this;
        }

        protected String getKey() {
            return key;
        }

        protected Object getO() {
            return o;
        }

        protected void setSst(Styles styles) {
            this.styles = styles;
        }

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
                throw new ExportException("Width " + width + " less than 0.");
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
        protected int getCellStyle(Class<?> clazz) {
            int style;
            if (isString(clazz)) {
                style = Styles.defaultStringBorderStyle();
            } else if (isDate(clazz) || isLocalDate(clazz)) {
                style = Styles.defaultDateBorderStyle();
            } else if (isDateTime(clazz) || isLocalDateTime(clazz)) {
                style = Styles.defaultTimestampBorderStyle();
            } else if (isBool(clazz) || isChar(clazz)) {
                style = Styles.clearHorizontal(Styles.defaultStringBorderStyle()) | Horizontals.CENTER;
            } else if (isInt(clazz) || isLong(clazz) || isBigDecimal(clazz)) {
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
            } else if (isFloat(clazz)) {
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
                style = styles.addNumFmt(tm);
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
     * 取消隔行变色
     * @return worksheet实例
     */
    public Sheet cancelOddStyle() {
        this.autoOdd = 1;
        return this;
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
     * 获取列头
     * @return 列头数组
     */
    public final Column[] getColumns() {
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
    public abstract void close();

    /**
     * write worksheet data to path
     * @param xl sheet.xml path
     * @throws IOException write error
     * @throws ExportException others
     */
    public abstract void writeTo(Path xl) throws IOException, ExportException;

    /**
     * 添加关联
     * @param rel Relationship
     * @return worksheet实例
     */
    Sheet addRel(Relationship rel) {
        relManager.add(rel);
        return this;
    }

    protected String getFileName() {
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

    private int defaultHeadStyle() {
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

    /**
     * 测试是否为Date类型
     * @param clazz 列类型
     * @return bool
     */
    static boolean isDate(Class<?> clazz) {
        return clazz == java.util.Date.class
                || clazz == java.sql.Date.class;
    }
    /**
     * 测试是否为DateTime类型
     * @param clazz 列类型
     * @return bool
     */
    static boolean isDateTime(Class<?> clazz) {
        return clazz == java.sql.Timestamp.class;
    }
    /**
     * 测试是否为Int类型
     * @param clazz 列类型
     * @return bool
     */
    static boolean isInt(Class<?> clazz) {
        return clazz == int.class || clazz == Integer.class
                || clazz == char.class || clazz == Character.class
                || clazz == byte.class || clazz == Byte.class
                || clazz == short.class || clazz == Short.class;
    }
    /**
     * 测试是否为Long类型
     * @param clazz 列类型
     * @return bool
     */
    static boolean isLong(Class<?> clazz) {
        return clazz == long.class || clazz == Long.class;
    }
    /**
     * 测试是否为Float类型
     * @param clazz 列类型
     * @return bool
     */
    static boolean isFloat(Class<?> clazz) {
        return clazz == double.class || clazz == Double.class
                || clazz == float.class || clazz == Float.class;
    }
    /**
     * 测试是否为Boolean类型
     * @param clazz 列类型
     * @return bool
     */
    static boolean isBool(Class<?> clazz) {
        return clazz == boolean.class || clazz == Boolean.class;
    }
    /**
     * 测试是否为String类型
     * @param clazz 列类型
     * @return bool
     */
    static boolean isString(Class<?> clazz) {
        return clazz == String.class || clazz == CharSequence.class;
    }
    /**
     * 测试是否为Char类型
     * @param clazz 列类型
     * @return bool
     */
    static boolean isChar(Class<?> clazz) {
        return clazz == char.class || clazz == Character.class;
    }
    /**
     * 测试是否为BigDecimal类型
     * @param clazz 列类型
     * @return bool
     */
    static boolean isBigDecimal(Class<?> clazz) {
        return clazz == BigDecimal.class;
    }
    /**
     * 测试是否为LocalDate类型
     * @param clazz 列类型
     * @return bool
     */
    static boolean isLocalDate(Class<?> clazz) {
        return clazz == LocalDate.class;
    }
    /**
     * 测试是否为LocalDateTime类型
     * @param clazz 列类型
     * @return bool
     */
    static boolean isLocalDateTime(Class<?> clazz) {
        return clazz == LocalDateTime.class;
    }
    /**
     * 测试是否为java.sql.Time类型
     * @param clazz 列类型
     * @return bool
     */
    static boolean isTime(Class<?> clazz) {
        return clazz == java.sql.Time.class;
    }
    /**
     * 测试是否为LocalTime类型
     * @param clazz 列类型
     * @return bool
     */
    static boolean isLocalTime(Class<?> clazz) {
        return clazz == java.time.LocalTime.class;
    }

    static boolean blockOrDefault(int style) {
        return style == -1 || style == Styles.defaultIntBorderStyle();
    }
    /**
     * 写worksheet头部
     * @param bw bufferedWriter
     */
    protected void writeBefore(ExtBufferedWriter bw) throws IOException {
        StringBuilder buf = new StringBuilder(Const.EXCEL_XML_DECLARATION);
        // Declaration
        buf.append(Const.lineSeparator); // new line
        // Root node
        if (getClass().isAnnotationPresent(TopNS.class)) {
            TopNS topNS = getClass().getAnnotation(TopNS.class);
            buf.append('<').append(topNS.value());
            String[] prefixs = topNS.prefix(), urls = topNS.uri();
            for (int i = 0, len = prefixs.length; i < len; ) {
                buf.append(" xmlns");
                if (prefixs[i] != null && !prefixs[i].isEmpty()) {
                    buf.append(':').append(prefixs[i]);
                }
                buf.append("=\"").append(urls[i]);
                if (++i < len) {
                    buf.append('"');
                }
            }
        } else {
            buf.append("<worksheet xmlns=\"").append(Const.SCHEMA_MAIN);
        }
        buf.append("\">");

        // Dimension
        buf.append("<dimension ref=\"A1\"/>");
        headInfoLen = buf.length() - 3;

        // SheetViews default value
        buf.append("<sheetViews><sheetView workbookViewId=\"0\"");
        if (id == 1) { // Default select the first worksheet
            buf.append(" tabSelected=\"1\"");
        }
        buf.append("/></sheetViews>");

        // Default format
        buf.append("<sheetFormatPr defaultRowHeight=\"16.5\" baseColWidth=\"");
        buf.append((int)width);
        buf.append("\"/>");

        baseInfoLen = buf.length() - headInfoLen;
        // Write base info
        bw.write(buf.toString());

        // Write body data
        bw.write("<sheetData>");

        // Write header
        int r = ++rows;
        bw.write("<row r=\"");
        bw.writeInt(r);
        bw.write("\" customHeight=\"1\" ht=\"18.6\" spans=\"1:"); // spans 指定row开始和结束行
        bw.writeInt(columns.length);
        bw.write("\">");

        int c = 1, defaultStyle = defaultHeadStyle();
        for (Column hc : columns) {
            bw.write("<c r=\"");
            bw.write(int2Col(c++));
            bw.writeInt(r);
            bw.write("\" t=\"inlineStr\" s=\"");
            bw.writeInt(defaultStyle);
            bw.write("\"><is><t>");
            bw.write(hc.getName());
            bw.write("</t></is></c>");
        }
        bw.write("</row>");
    }

    /**
     * 写尾部
     * @param bw bufferedWriter
     */
    protected void writeAfter(ExtBufferedWriter bw) throws IOException {
        // End target --sheetData
        bw.write("</sheetData>");

        // background image
        if (waterMark != null) {
            // relationship
            Relationship r = relManager.likeByTarget("media/image"); // only one background image
            if (r != null) {
                bw.write("<picture r:id=\"");
                bw.write(r.getId());
                bw.write("\"/>");
            }
        }
        // End target
        if (getClass().isAnnotationPresent(TopNS.class)) {
            TopNS topNS = getClass().getAnnotation(TopNS.class);
            bw.write("</");
            bw.write(topNS.value());
            bw.write('>');
        } else {
            bw.write("</worksheet>");
        }
        workbook.what("0009", getName(), String.valueOf(rows));
    }

    /**
     * 写行数据
     * @param rs ResultSet
     * @param bw bufferedWriter
     */
    protected void writeRow(ResultSet rs, ExtBufferedWriter bw) throws IOException, SQLException {
        // Row number
        int r = ++rows;
        // logging
        if (r % 1_0000 == 0) {
            workbook.what("0014", String.valueOf(r));
        }
        final int len = columns.length;
        bw.write("<row r=\"");
        bw.writeInt(r);
        // default data row height 16.5
        bw.write("\" spans=\"1:");
        bw.writeInt(len);
        bw.write("\">");

        for (int i = 0; i < len; i++) {
            Column hc = columns[i];

            // t n=numeric (default), s=string, b=boolean, str=function string
            // TODO function <f ca="1" or t="shared" ref="O10:O15" si="0" ... si="10"></f>
            if (isString(hc.clazz)) {
                String s = rs.getString(i + 1);
                writeString(bw, s, i);
            }
            else if (isDate(hc.clazz)) {
                java.sql.Date date = rs.getDate(i + 1);
                writeDate(bw, date, i);
            }
            else if (isDateTime(hc.clazz)) {
                Timestamp ts = rs.getTimestamp(i + 1);
                writeTimestamp(bw, ts, i);
            }
//            else if (isChar(hc.clazz)) {
//                char c = (char) rs.getInt(i + 1);
//                writeChar(bw, c, i);
//            }
            else if (isInt(hc.clazz)) {
                int n = rs.getInt(i + 1);
                writeInt(bw, n, i);
            }
            else if (isLong(hc.clazz)) {
                long l = rs.getLong(i + 1);
                writeLong(bw, l, i);
            }
            else if (isFloat(hc.clazz)) {
                double d = rs.getDouble(i + 1);
                writeDouble(bw, d, i);
            } else if (isBool(hc.clazz)) {
                boolean bool = rs.getBoolean(i + 1);
                writeBoolean(bw, bool, i);
            } else if (isBigDecimal(hc.clazz)) {
                writeBigDecimal(bw, rs.getBigDecimal(i + 1), i);
            } else if (isTime(hc.clazz)) {
                writeTime(bw, rs.getTime(i + 1), i);
            } else {
                Object o = rs.getObject(i + 1);
                if (o != null) {
                    writeString(bw, o.toString(), i);
                } else {
                    writeNull(bw, i);
                }
            }
        }
        bw.write("</row>");
    }

    /**
     * 写行数据
     * @param rs ResultSet
     * @param bw
     */
    protected void writeRowAutoSize(ResultSet rs, ExtBufferedWriter bw) throws IOException, SQLException {
        int r = ++rows;
        // logging
        if (r % 1_0000 == 0) {
            workbook.what("0014", String.valueOf(r));
        }
        final int len = columns.length;
        bw.write("<row r=\"");
        bw.writeInt(r);
        bw.write("\" spans=\"1:");
        bw.writeInt(len);
        bw.write("\">");

        for (int i = 0; i < len; i++) {
            Column hc = columns[i];
            // t n=numeric (default), s=string, b=boolean, str=function string
            // TODO function <f ca="1" or t="shared" ref="O10:O15" si="0" ... si="10"></f>
            if (isString(hc.clazz)) {
                String s = rs.getString(i + 1);
                writeStringAutoSize(bw, s, i);
            }
            else if (isDate(hc.clazz)) {
                java.sql.Date date = rs.getDate(i + 1);
                writeDate(bw, date, i);
            }
            else if (isDateTime(hc.clazz)) {
                Timestamp ts = rs.getTimestamp(i + 1);
                writeTimestamp(bw, ts, i);
            }
            else if (isInt(hc.clazz)) {
                int n = rs.getInt(i + 1);
                writeIntAutoSize(bw, n, i);
            }
            else if (isLong(hc.clazz)) {
                long l = rs.getLong(i + 1);
                writeLong(bw, l, i);
            }
            else if (isFloat(hc.clazz)) {
                double d = rs.getDouble(i + 1);
                writeDouble(bw, d, i);
            } else if (isBool(hc.clazz)) {
                boolean bool = rs.getBoolean(i + 1);
                writeBoolean(bw, bool, i);
            } else if (isBigDecimal(hc.clazz)) {
                writeBigDecimal(bw, rs.getBigDecimal(i + 1), i);
            } else if (isTime(hc.clazz)) {
                writeTime(bw, rs.getTime(i + 1), i);
            } else {
                Object o = rs.getObject(i + 1);
                if (o != null) {
                    writeStringAutoSize(bw, o.toString(), i);
                } else {
                    writeNull(bw, i);
                }
            }
        }
        bw.write("</row>");
    }

    protected int getStyleIndex(Column hc, Object o) {
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
     * 写字符串
     * @throws IOException
     */
    protected void writeString(ExtBufferedWriter bw, String s, int column) throws IOException {
        writeString(bw, s, column, s);
    }

    private void writeString(ExtBufferedWriter bw, String s, int column, Object o) throws IOException {
        Column hc = columns[column];
        int styleIndex = getStyleIndex(hc, o);
        bw.write("<c r=\"");
        bw.write(int2Col(column + 1));
        bw.writeInt(rows);
        int i;
        if (StringUtil.isEmpty(s)) {
            bw.write("\" s=\"");
            bw.writeInt(styleIndex);
            bw.write("\"/>");
        }
        else if (hc.isShare() && (i = workbook.getSst().get(s)) >= 0) {
            bw.write("\" t=\"s\" s=\"");
            bw.writeInt(styleIndex);
            bw.write("\"><v>");
            bw.writeInt(i);
            bw.write("</v></c>");
        }
        else {
            bw.write("\" t=\"inlineStr\" s=\"");
            bw.writeInt(styleIndex);
            bw.write("\"><is><t>");
            bw.escapeWrite(s); // escape text
            bw.write("</t></is></c>");
        }
    }

    protected void writeStringAutoSize(ExtBufferedWriter bw, String s, int column) throws IOException {
        writeStringAutoSize(bw, s, column, s);
    }

    protected void writeStringAutoSize(ExtBufferedWriter bw, String s, int column, Object o) throws IOException {
        Column hc = columns[column];
        int styleIndex = getStyleIndex(hc, o);
        bw.write("<c r=\"");
        bw.write(int2Col(column + 1));
        bw.writeInt(rows);
        if (StringUtil.isEmpty(s)) {
            bw.write("\" s=\"");
            bw.writeInt(styleIndex);
            bw.write("\"/>");
        } else {
            int i;
            if (hc.isShare() && (i = workbook.getSst().get(s)) >= 0) {
                bw.write("\" t=\"s\" s=\"");
                bw.writeInt(styleIndex);
                bw.write("\"><v>");
                bw.writeInt(i);
                bw.write("</v></c>");
            } else {
                bw.write("\" t=\"inlineStr\" s=\"");
                bw.writeInt(styleIndex);
                bw.write("\"><is><t>");
                bw.escapeWrite(s); // escape text
                bw.write("</t></is></c>");
            }
            int ln = s.getBytes("GB2312").length; // TODO 计算
            if (hc.width == 0 && (hc.o == null || (int) hc.o < ln)) {
                hc.o = ln;
            }
        }
    }

    protected void writeDate(ExtBufferedWriter bw, Date date, int column) throws IOException {
        writeDate(bw, date, column, date);
    }

    protected void writeDate(ExtBufferedWriter bw, Date date, int column, Object o) throws IOException {
        int styleIndex = getStyleIndex(columns[column], o);
        bw.write("<c r=\"");
        bw.write(int2Col(column + 1));
        bw.writeInt(rows);
        if (date == null) {
            bw.write("\" s=\"");
            bw.writeInt(styleIndex);
            bw.write("\"/>");
        } else {
            bw.write("\" s=\"");
            bw.writeInt(styleIndex);
            bw.write("\"><v>");
            bw.writeInt(toDateValue(date));
            bw.write("</v></c>");
        }
    }

    protected void writeLocalDate(ExtBufferedWriter bw, LocalDate date, int column) throws IOException {
        writeLocalDate(bw, date, column, date);
    }

    protected void writeLocalDate(ExtBufferedWriter bw, LocalDate date, int column, Object o) throws IOException {
        int styleIndex = getStyleIndex(columns[column], o);
        bw.write("<c r=\"");
        bw.write(int2Col(column + 1));
        bw.writeInt(rows);
        if (date == null) {
            bw.write("\" s=\"");
            bw.writeInt(styleIndex);
            bw.write("\"/>");
        } else {
            bw.write("\" s=\"");
            bw.writeInt(styleIndex);
            bw.write("\"><v>");
            bw.writeInt(toDateValue(date));
            bw.write("</v></c>");
        }
    }

    protected void writeTimestamp(ExtBufferedWriter bw, Timestamp ts, int column) throws IOException {
        writeTimestamp(bw, ts, column, ts);
    }

    protected void writeTimestamp(ExtBufferedWriter bw, Timestamp ts, int column, Object o) throws IOException {
        int styleIndex = getStyleIndex(columns[column], o);
        bw.write("<c r=\"");
        bw.write(int2Col(column + 1));
        bw.writeInt(rows);
        if (ts == null) {
            bw.write("\" s=\"");
            bw.writeInt(styleIndex);
            bw.write("\"/>");
        } else {
            bw.write("\" s=\"");
            bw.writeInt(styleIndex);
            bw.write("\"><v>");
            bw.write(toDateTimeValue(ts));
            bw.write("</v></c>");
        }
    }

    protected void writeLocalDateTime(ExtBufferedWriter bw, LocalDateTime ts, int column) throws IOException {
        writeLocalDateTime(bw, ts, column, ts);
    }

    protected void writeLocalDateTime(ExtBufferedWriter bw, LocalDateTime ts, int column, Object o) throws IOException {
        int styleIndex = getStyleIndex(columns[column], o);
        bw.write("<c r=\"");
        bw.write(int2Col(column + 1));
        bw.writeInt(rows);
        if (ts == null) {
            bw.write("\" s=\"");
            bw.writeInt(styleIndex);
            bw.write("\"/>");
        } else {
            bw.write("\" s=\"");
            bw.writeInt(styleIndex);
            bw.write("\"><v>");
            bw.write(toDateTimeValue(ts));
            bw.write("</v></c>");
        }
    }

    protected void writeTime(ExtBufferedWriter bw, Time date, int column) throws IOException {
        writeTime(bw, date, column, date);
    }

    protected void writeTime(ExtBufferedWriter bw, Time date, int column, Object o) throws IOException {
        int styleIndex = getStyleIndex(columns[column], o);
        bw.write("<c r=\"");
        bw.write(int2Col(column + 1));
        bw.writeInt(rows);
        if (date == null) {
            bw.write("\" s=\"");
            bw.writeInt(styleIndex);
            bw.write("\"/>");
        } else {
            bw.write("\" s=\"");
            bw.writeInt(styleIndex);
            bw.write("\"><v>");
            bw.write(toTimeValue(date));
            bw.write("</v></c>");
        }
    }

    protected void writeLocalTime(ExtBufferedWriter bw, LocalTime date, int column) throws IOException {
        writeLocalTime(bw, date, column, date);
    }

    protected void writeLocalTime(ExtBufferedWriter bw, LocalTime date, int column, Object o) throws IOException {
        int styleIndex = getStyleIndex(columns[column], o);
        bw.write("<c r=\"");
        bw.write(int2Col(column + 1));
        bw.writeInt(rows);
        if (date == null) {
            bw.write("\" s=\"");
            bw.writeInt(styleIndex);
            bw.write("\"/>");
        } else {
            bw.write("\" s=\"");
            bw.writeInt(styleIndex);
            bw.write("\"><v>");
            bw.write(toTimeValue(date));
            bw.write("</v></c>");
        }
    }

    protected void writeBigDecimal(ExtBufferedWriter bw, BigDecimal bd, int column) throws IOException {
        writeBigDecimal(bw, bd, column, bd);
    }

    protected void writeBigDecimal(ExtBufferedWriter bw, BigDecimal bd, int column, Object o) throws IOException {
        int styleIndex = getStyleIndex(columns[column], o);
        bw.write("<c r=\"");
        bw.write(int2Col(column + 1));
        bw.writeInt(rows);
        if (bd == null) {
            bw.write("\" s=\"");
            bw.writeInt(styleIndex);
            bw.write("\"/>");
        } else {
            bw.write("\" s=\"");
            bw.writeInt(styleIndex);
            bw.write("\"><v>");
            bw.write(bd.toString());
            bw.write("</v></c>");
        }
    }

    protected void writeInt(ExtBufferedWriter bw, int n, int column) throws IOException {
        Column hc = columns[column];
        if (hc.processor == null) {
            writeInt0(bw, n, column);
        } else {
            Object o = hc.processor.conversion(n);
            if (o != null) {
                Class<?> clazz = o.getClass();
                boolean blockOrDefault = blockOrDefault(hc.cellStyle);
                if (isString(clazz)) {
                    if (blockOrDefault) {
                        hc.cellStyle = hc.getCellStyle(String.class);
                    }
                    writeString(bw, o.toString(), column, n);
                }
                else if (isChar(clazz)) {
                    if (blockOrDefault) {
                        hc.cellStyle = Styles.defaultCharBorderStyle();
                    }
                    char c = ((Character) o).charValue();
                    writeChar0(bw, c, column, n);
                }
                else if (isInt(clazz)) {
                    n = ((Integer) o).intValue();
                    writeInt0(bw, n, column, n);
                }
                else if (isLong(clazz)) {
                    long l = ((Long) o).longValue();
                    writeLong(bw, l, column, n);
                }
                else if (isDate(clazz)) {
                    if (blockOrDefault) {
                        hc.cellStyle = hc.getCellStyle(java.util.Date.class);
                    }
                    writeDate(bw, (Date) o, column, n);
                }
                else if (isDateTime(clazz)) {
                    if (blockOrDefault) {
                        hc.cellStyle = hc.getCellStyle(java.sql.Timestamp.class);
                    }
                    writeTimestamp(bw, (Timestamp) o, column, n);
                }
                else if (isFloat(clazz)) {
                    if (blockOrDefault) {
                        hc.cellStyle = hc.getCellStyle(double.class);
                    }
                    writeDouble(bw, ((Double) o).doubleValue(), column, n);
                }
                else if (isBool(clazz)) {
                    if (blockOrDefault) {
                        hc.cellStyle = hc.getCellStyle(boolean.class);
                    }
                    boolean bool = ((Boolean) o).booleanValue();
                    writeBoolean(bw, bool, column, n);
                }
                else if (isBigDecimal(clazz)) {
                    if (blockOrDefault) {
                        hc.cellStyle = hc.getCellStyle(BigDecimal.class);
                    }
                    writeBigDecimal(bw, (BigDecimal) o, column, n);
                }
                else if (isTime(clazz)) {
                    if (blockOrDefault) {
                        hc.cellStyle = hc.getCellStyle(Time.class);
                    }
                    writeTime(bw, (Time) o, column, n);
                }
                else  if (isLocalDate(clazz)) {
                    if (blockOrDefault) {
                        hc.cellStyle = hc.getCellStyle(LocalDate.class);
                    }
                    writeLocalDate(bw, (LocalDate) o, column, n);
                }
                else  if (isLocalDateTime(clazz)) {
                    if (blockOrDefault) {
                        hc.cellStyle = hc.getCellStyle(LocalDateTime.class);
                    }
                    writeLocalDateTime(bw, (LocalDateTime) o, column, n);
                }
                else  if (isLocalTime(clazz)) {
                    if (blockOrDefault) {
                        hc.cellStyle = hc.getCellStyle(LocalTime.class);
                    }
                    writeLocalTime(bw, (LocalTime) o, column, n);
                }
                else {
                    if (blockOrDefault) {
                        hc.cellStyle = hc.getCellStyle(String.class);
                    }
                    writeString(bw, o.toString(), column, n);
                }
            }
            else {
                writeNull(bw, column);
            }
        }
    }

    protected void writeIntAutoSize(ExtBufferedWriter bw, int n, int column) throws IOException {
        Column hc = columns[column];
        if (hc.processor == null) {
            writeInt0(bw, n, column);
        } else {
            Object o = hc.processor.conversion(n);
            if (o != null) {
                Class<?> clazz = o.getClass();
                boolean blockOrDefault = blockOrDefault(hc.cellStyle);
                if (isString(clazz)) {
                    if (blockOrDefault) {
                        hc.cellStyle = hc.getCellStyle(String.class);
                    }
                    writeStringAutoSize(bw, o.toString(), column, n);
                }
                else if (isChar(clazz)) {
                    if (blockOrDefault) {
                        hc.cellStyle = Styles.defaultCharBorderStyle();
                    }
                    char c = ((Character) o).charValue();
                    writeChar0(bw, c, column, n);
                }
                else if (isInt(clazz)) {
                    int nn = ((Integer) o).intValue();
                    writeInt0(bw, nn, column, n);
                }
                else if (isLong(clazz)) {
                    long l = ((Long) o).longValue();
                    writeLong(bw, l, column, n);
                }
                else if (isDate(clazz)) {
                    if (blockOrDefault) {
                        hc.cellStyle = hc.getCellStyle(java.util.Date.class);
                    }
                    writeDate(bw, (Date) o, column, n);
                }
                else if (isDateTime(clazz)) {
                    if (blockOrDefault) {
                        hc.cellStyle = hc.getCellStyle(java.sql.Timestamp.class);
                    }
                    writeTimestamp(bw, (Timestamp) o, column, n);
                }
                else if (isFloat(clazz)) {
                    if (blockOrDefault) {
                        hc.cellStyle = hc.getCellStyle(double.class);
                    }
                    writeDouble(bw, ((Double) o).doubleValue(), column, n);
                }
                else if (isBool(clazz)) {
                    if (blockOrDefault) {
                        hc.cellStyle = hc.getCellStyle(boolean.class);
                    }
                    boolean bool = ((Boolean) o).booleanValue();
                    writeBoolean(bw, bool, column, n);
                }
                else if (isBigDecimal(clazz)) {
                    if (blockOrDefault) {
                        hc.cellStyle = hc.getCellStyle(BigDecimal.class);
                    }
                    writeBigDecimal(bw, (BigDecimal) o, column, n);
                }
                else if (isTime(clazz)) {
                    if (blockOrDefault) {
                        hc.cellStyle = hc.getCellStyle(Time.class);
                    }
                    writeTime(bw, (Time) o, column, n);
                }
                else  if (isLocalDate(clazz)) {
                    if (blockOrDefault) {
                        hc.cellStyle = hc.getCellStyle(LocalDate.class);
                    }
                    writeLocalDate(bw, (LocalDate) o, column, n);
                }
                else  if (isLocalDateTime(clazz)) {
                    if (blockOrDefault) {
                        hc.cellStyle = hc.getCellStyle(LocalDateTime.class);
                    }
                    writeLocalDateTime(bw, (LocalDateTime) o, column, n);
                }
                else  if (isLocalTime(clazz)) {
                    if (blockOrDefault) {
                        hc.cellStyle = hc.getCellStyle(LocalTime.class);
                    }
                    writeLocalTime(bw, (LocalTime) o, column, n);
                }
                else {
                    if (blockOrDefault) {
                        hc.cellStyle = hc.getCellStyle(String.class);
                    }
                    writeStringAutoSize(bw, o.toString(), column, n);
                }
            }
            else {
                writeNull(bw, column);
            }
        }
    }

    protected void writeChar(ExtBufferedWriter bw, char c, int column) throws IOException {
        Column hc = columns[column];
        if (hc.processor == null) {
            writeChar0(bw, c, column);
        } else {
            Object o = hc.processor.conversion(c);
            if (o != null) {
                Class<?> clazz = o.getClass();
                boolean blockOrDefault = blockOrDefault(hc.cellStyle);
                if (isString(clazz)) {
                    if (blockOrDefault) {
                        hc.cellStyle = hc.getCellStyle(String.class);
                    }
                    writeString(bw, o.toString(), column, c);
                }
                else if (isChar(clazz)) {
                    char cc = ((Character) o).charValue();
                    writeChar0(bw, cc, column, c);
                }
                else if (isInt(clazz)) {
                    if (blockOrDefault) {
                        hc.cellStyle = hc.getCellStyle(int.class);
                    }
                    int n = ((Integer) o).intValue();
                    writeInt0(bw, n, column, c);
                }
                else if (isLong(clazz)) {
                    if (blockOrDefault) {
                        hc.cellStyle = hc.getCellStyle(long.class);
                    }
                    long l = ((Long) o).longValue();
                    writeLong(bw, l, column, c);
                }
                else if (isDate(clazz)) {
                    if (blockOrDefault) {
                        hc.cellStyle = hc.getCellStyle(java.util.Date.class);
                    }
                    writeDate(bw, (Date) o, column, c);
                }
                else if (isDateTime(clazz)) {
                    if (blockOrDefault) {
                        hc.cellStyle = hc.getCellStyle(java.sql.Timestamp.class);
                    }
                    writeTimestamp(bw, (Timestamp) o, column, c);
                }
                else if (isFloat(clazz)) {
                    if (blockOrDefault) {
                        hc.cellStyle = hc.getCellStyle(double.class);
                    }
                    writeDouble(bw, ((Double) o).doubleValue(), column, c);
                }
                else if (isBool(clazz)) {
                    if (blockOrDefault) {
                        hc.cellStyle = hc.getCellStyle(boolean.class);
                    }
                    boolean bool = ((Boolean) o).booleanValue();
                    writeBoolean(bw, bool, column, c);
                }
                else if (isBigDecimal(clazz)) {
                    if (blockOrDefault) {
                        hc.cellStyle = hc.getCellStyle(BigDecimal.class);
                    }
                    writeBigDecimal(bw, (BigDecimal) o, column, c);
                }
                else if (isTime(clazz)) {
                    if (blockOrDefault) {
                        hc.cellStyle = hc.getCellStyle(Time.class);
                    }
                    writeTime(bw, (Time) o, column, c);
                }
                else  if (isLocalDate(clazz)) {
                    if (blockOrDefault) {
                        hc.cellStyle = hc.getCellStyle(LocalDate.class);
                    }
                    writeLocalDate(bw, (LocalDate) o, column, c);
                }
                else  if (isLocalDateTime(clazz)) {
                    if (blockOrDefault) {
                        hc.cellStyle = hc.getCellStyle(LocalDateTime.class);
                    }
                    writeLocalDateTime(bw, (LocalDateTime) o, column, c);
                }
                else  if (isLocalTime(clazz)) {
                    if (blockOrDefault) {
                        hc.cellStyle = hc.getCellStyle(LocalTime.class);
                    }
                    writeLocalTime(bw, (LocalTime) o, column, c);
                }
                else {
                    if (blockOrDefault) {
                        hc.cellStyle = hc.getCellStyle(String.class);
                    }
                    writeString(bw, o.toString(), column, c);
                }
            }
            else {
                writeNull(bw, column);
            }
        }
    }

    protected void writeCharAutoSize(ExtBufferedWriter bw, char c, int column) throws IOException {
        Column hc = columns[column];
        if (hc.processor == null) {
            writeChar0(bw, c, column);
        } else {
            Object o = hc.processor.conversion(c);
            if (o != null) {
                Class<?> clazz = o.getClass();
                boolean blockOrDefault = blockOrDefault(hc.cellStyle);
                if (isString(clazz)) {
                    if (blockOrDefault) {
                        hc.cellStyle = hc.getCellStyle(String.class);
                    }
                    writeStringAutoSize(bw, o.toString(), column);
                }
                else if (isChar(clazz)) {
                    char cc = ((Character) o).charValue();
                    writeChar0(bw, cc, column, c);
                }
                else if (isInt(clazz)) {
                    if (blockOrDefault) {
                        hc.cellStyle = hc.getCellStyle(int.class);
                    }
                    int n = ((Integer) o).intValue();
                    writeInt0(bw, n, column, c);
                }
                else if (isLong(clazz)) {
                    if (blockOrDefault) {
                        hc.cellStyle = hc.getCellStyle(long.class);
                    }
                    long l = ((Long) o).longValue();
                    writeLong(bw, l, column, c);
                }
                else if (isDate(clazz)) {
                    if (blockOrDefault) {
                        hc.cellStyle = hc.getCellStyle(java.util.Date.class);
                    }
                    writeDate(bw, (Date) o, column, c);
                }
                else if (isDateTime(clazz)) {
                    if (blockOrDefault) {
                        hc.cellStyle = hc.getCellStyle(java.sql.Timestamp.class);
                    }
                    writeTimestamp(bw, (Timestamp) o, column, c);
                }
                else if (isFloat(clazz)) {
                    if (blockOrDefault) {
                        hc.cellStyle = hc.getCellStyle(double.class);
                    }
                    writeDouble(bw, ((Double) o).doubleValue(), column, c);
                }
                else if (isBool(clazz)) {
                    if (blockOrDefault) {
                        hc.cellStyle = hc.getCellStyle(boolean.class);
                    }
                    boolean bool = ((Boolean) o).booleanValue();
                    writeBoolean(bw, bool, column, c);
                }
                else if (isBigDecimal(clazz)) {
                    if (blockOrDefault) {
                        hc.cellStyle = hc.getCellStyle(BigDecimal.class);
                    }
                    writeBigDecimal(bw, (BigDecimal) o, column, c);
                }
                else if (isTime(clazz)) {
                    if (blockOrDefault) {
                        hc.cellStyle = hc.getCellStyle(Time.class);
                    }
                    writeTime(bw, (Time) o, column, c);
                }
                else  if (isLocalDate(clazz)) {
                    if (blockOrDefault) {
                        hc.cellStyle = hc.getCellStyle(LocalDate.class);
                    }
                    writeLocalDate(bw, (LocalDate) o, column, c);
                }
                else  if (isLocalDateTime(clazz)) {
                    if (blockOrDefault) {
                        hc.cellStyle = hc.getCellStyle(LocalDateTime.class);
                    }
                    writeLocalDateTime(bw, (LocalDateTime) o, column, c);
                }
                else  if (isLocalTime(clazz)) {
                    if (blockOrDefault) {
                        hc.cellStyle = hc.getCellStyle(LocalTime.class);
                    }
                    writeLocalTime(bw, (LocalTime) o, column, c);
                }
                else {
                    if (blockOrDefault) {
                        hc.cellStyle = hc.getCellStyle(String.class);
                    }
                    writeStringAutoSize(bw, o.toString(), column, c);
                }
            } else {
                writeNull(bw, column);
            }
        }
    }
    private void writeInt0(ExtBufferedWriter bw, int n, int column) throws IOException {
        writeInt0(bw, n, column, n);
    }

    private void writeInt0(ExtBufferedWriter bw, int n, int column, Object o) throws IOException {
        int styleIndex = getStyleIndex(columns[column], o);
        bw.write("<c r=\"");
        bw.write(int2Col(column + 1));
        bw.writeInt(rows);
        bw.write("\" s=\"");
        bw.writeInt(styleIndex);
        bw.write("\"><v>");
        bw.writeInt(n);
        bw.write("</v></c>");
    }

    private void writeChar0(ExtBufferedWriter bw, char c, int column) throws IOException {
        writeChar0(bw, c, column, c);
    }

    private void writeChar0(ExtBufferedWriter bw, char c, int column, Object o) throws IOException {
        int styleIndex = getStyleIndex(columns[column], o);
        bw.write("<c r=\"");
        bw.write(int2Col(column + 1));
        bw.writeInt(rows);
        bw.write("\" t=\"s\" s=\"");
        bw.writeInt(styleIndex);
        bw.write("\"><v>");
        bw.writeInt(workbook.getSst().get(c));
        bw.write("</v></c>");
    }

    protected void writeLong(ExtBufferedWriter bw, long l, int column) throws IOException {
        writeLong(bw, l, column, l);
    }

    protected void writeLong(ExtBufferedWriter bw, long l, int column, Object o) throws IOException {
        int styleIndex = getStyleIndex(columns[column], o);
        bw.write("<c r=\"");
        bw.write(int2Col(column + 1));
        bw.writeInt(rows);
        bw.write("\" s=\"");
        bw.writeInt(styleIndex);
        bw.write("\"><v>");
        bw.write(l);
        bw.write("</v></c>");
    }

    protected void writeDouble(ExtBufferedWriter bw, double d, int column) throws IOException {
        writeDouble(bw, d, column, d);
    }

    protected void writeDouble(ExtBufferedWriter bw, double d, int column, Object o) throws IOException {
        int styleIndex = getStyleIndex(columns[column], o);
        bw.write("<c r=\"");
        bw.write(int2Col(column + 1));
        bw.writeInt(rows);
        bw.write("\" s=\"");
        bw.writeInt(styleIndex);
        bw.write("\"><v>");
        bw.write(d);
        bw.write("</v></c>");
    }

    protected void writeBoolean(ExtBufferedWriter bw, boolean bool, int column) throws IOException {
        writeBoolean(bw, bool, column, bool);
    }

    protected void writeBoolean(ExtBufferedWriter bw, boolean bool, int column, Object o) throws IOException {
        int styleIndex = getStyleIndex(columns[column], o);
        bw.write("<c r=\"");
        bw.write(int2Col(column + 1));
        bw.writeInt(rows);
        bw.write("\" t=\"b\" s=\"");
        bw.writeInt(styleIndex);
        bw.write("\"><v>");
        bw.writeInt(bool ? 1 : 0);
        bw.write("</v></c>");
    }

    protected void writeNull(ExtBufferedWriter bw, int column) throws IOException {
        int styleIndex = getStyleIndex(columns[column], null);
        bw.write("<c r=\"");
        bw.write(int2Col(column + 1));
        bw.writeInt(rows);
        bw.write("\" s=\"");
        bw.writeInt(styleIndex);
        bw.write("\"/>");
    }

    /**
     * 写空行数据
     * @param bw
     */
    protected void writeEmptyRow(ExtBufferedWriter bw) throws IOException {
        // Row number
        int r = ++rows;
        final int len = columns.length;
        bw.write("<row r=\"");
        bw.writeInt(r);
        bw.write("\" ht=\"16.5\" spans=\"1:");
        bw.writeInt(len);
        bw.write("\">");

        Styles styles = workbook.getStyles();
        for (int i = 1; i <= len; i++) {
            Column hc = columns[i - 1];
            bw.write("<c r=\"");
            bw.write(int2Col(i));
            bw.writeInt(r);

            int style = hc.getCellStyle();
            // 隔行变色
            if (autoOdd == 0 && isOdd() && !Styles.hasFill(style)) {
                style |= oddFill;
            }
            int styleIndex = styles.of(style);
            bw.write("\" s=\"");
            bw.writeInt(styleIndex);
            bw.write("\"/>");

            if (hc.getO() == null) {
                hc.setO(hc.getName().getBytes("GB2312").length);
            }
        }
        bw.write("</row>");
    }

    protected  void autoColumnSize(File sheet) throws IOException {
        // resize each column width ...
        File temp = new File(sheet.getParent(), sheet.getName() + ".temp");
        if (!sheet.renameTo(temp)) {
            Files.move(sheet.toPath(), temp.toPath(), StandardCopyOption.REPLACE_EXISTING);
        }

        FileChannel inChannel = null;
        FileChannel outChannel = null;
        try (FileInputStream fis = new FileInputStream(temp);
        		FileOutputStream fos = new FileOutputStream(sheet)) {
        	inChannel = fis.getChannel();
            outChannel = fos.getChannel();
            
            inChannel.transferTo(0, headInfoLen, outChannel);
            ByteBuffer buffer = ByteBuffer.allocate(baseInfoLen);
            inChannel.read(buffer, headInfoLen);
            buffer.compact();
            byte b;
            if ((b = buffer.get()) == '"') {
                char[] chars = int2Col(columns.length);
                String s = ':' + new String(chars) + rows;
                outChannel.write(ByteBuffer.wrap(s.getBytes(Const.UTF_8)));
            }
            buffer.flip();
            buffer.put(b);
            buffer.compact();
            outChannel.write(buffer);

            StringBuilder buf = new StringBuilder();
            buf.append("<cols>");
            int i = 0;
            for (Column hc : columns) {
                i++;
                buf.append("<col customWidth=\"1\" width=\"");
                if (hc.width > 0.0000001) {
                    buf.append(hc.width);
                    buf.append("\" max=\"");
                    buf.append(i);
                    buf.append("\" min=\"");
                    buf.append(i);
                    buf.append("\"/>");
                } else if (autoSize == 1) {
                    int _l = hc.name.getBytes("GB2312").length, len;
                    // TODO 根据字体字号计算文本宽度
                    if (isString(hc.clazz)) {
                        if (hc.o == null) {
                            len = 0;
                        } else {
                            len = (int) hc.o;
                        }
//                        len = hc.o.toString().getBytes("GB2312").length;
                    }
                    else if (isDate(hc.clazz) || isLocalDate(hc.clazz)) {
                        len = 10;
                    }
                    else if (isDateTime(hc.clazz) || isLocalDateTime(hc.clazz)) {
                        len = 20;
                    }
                    else if (isInt(hc.clazz)) {
                        // TODO 根据numFmt计算字符宽度
                        len = hc.type > 0 ? 12 :  11;
                    }
                    else if (isFloat(hc.clazz)) {
                        // TODO 根据numFmt计算字符宽度
                        if (hc.o == null) {
                            len = 0;
                        } else {
                            len = hc.o.toString().getBytes("GB2312").length;
                        }
//                        if (len < 11) {
//                            len = hc.type > 0 ? 12 : 11;
//                        }
                    } else if (isTime(hc.clazz) || isLocalTime(hc.clazz)) {
                        len = 8;
                    } else {
                        len = 10;
                    }
                    buf.append(_l > len ? _l + 3.38 : len + 3.38);
                    buf.append("\" max=\"");
                    buf.append(i);
                    buf.append("\" min=\"");
                    buf.append(i);
                    buf.append("\" bestFit=\"1\"/>");
                } else {
                    buf.append(width);
                    buf.append("\" max=\"");
                    buf.append(i);
                    buf.append("\" min=\"");
                    buf.append(i);
                    buf.append("\"/>");
                }
            }
            buf.append("</cols>");

            outChannel.write(ByteBuffer.wrap(buf.toString().getBytes(Const.UTF_8)));
            int start = headInfoLen + baseInfoLen;
            inChannel.transferTo(start, inChannel.size() - start, outChannel);

        } catch (IOException e) {
            throw e;
        } finally {
            boolean delete = temp.delete();
            if (!delete) {
                workbook.what("9005", temp.getAbsolutePath());
            }
            if (inChannel != null) {
            	inChannel.close();
            }
            if (outChannel != null) {
            	outChannel.close();
            }
        }
    }

    /**
     * Int转列号A-Z
     */
    private ThreadLocal<char[][]> cache = ThreadLocal.withInitial(() -> new char[][] {{65}, {65, 65}, {65, 65, 65}});
    private char[] int2Col(int n) {
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

}

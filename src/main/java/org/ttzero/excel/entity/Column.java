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

import org.ttzero.excel.drawing.Effect;
import org.ttzero.excel.entity.style.Border;
import org.ttzero.excel.entity.style.Fill;
import org.ttzero.excel.entity.style.Font;
import org.ttzero.excel.entity.style.Horizontals;
import org.ttzero.excel.entity.style.NumFmt;
import org.ttzero.excel.entity.style.Styles;
import org.ttzero.excel.entity.style.Verticals;
import org.ttzero.excel.manager.Const;
import org.ttzero.excel.processor.ConversionProcessor;
import org.ttzero.excel.processor.Converter;
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
 * Excel列，{@code Column}用于收集列属性将Java实体与Excel列进行映射，
 * 多个{@code Column}组成Excel表头行，当前最大支持1024层表头
 *
 * @author guanquan.wang at 2021-08-29 19:49
 */
public class Column {
    /**
     * 未实例化的列，可用于在写超出预知范围外的列
     */
    public static final Column UNALLOCATED_COLUMN = new Column();
    /**
     * Java对象中的字段、Map的Key或者SQL语句中的select字段，具体值由工作表类型决定
     */
    public String key;
    /**
     * Excel表头名称，未指定表头名时默认以{@code #key}的值代替
     */
    public String name;
    /**
     * Excel列值的类型，不特殊指定时该类型与Java对象类型一致，它影响最终输出到Excel单元格值的类型和对齐方式
     * 默认情况下文本类型左对齐，数字右对齐，日期居中，表头单元格全部居中
     */
    public Class<?> clazz;
    /**
     * 输出转换器，通常用于将不可读的状态值或枚举值转换为可读的文本输出到Excel
     */
    public ConversionProcessor processor;
    /**
     * 数据转换器，与{@code ConversionProcessor}不同的是这是一个双向转换器，
     * 同时{@code Converter}继承{@code ConversionProcessor}接口，当{@code processor}与
     * {@code converter}同时存在时前者具有更高的优先级
     */
    public Converter<?> converter;
    /**
     * 动态样式转换器，可根据单元格的值动态设置单元格或整行样式，通常用于高亮显示某些需要重视的行或单元格
     */
    public StyleProcessor styleProcessor;
    /**
     * 表格体的样式值
     */
    public Integer cellStyle;
    /**
     * 表格头的样式值
     */
    public Integer headerStyle;
    /**
     * 表格体的样式索引, -1表示未设置
     */
    protected int cellStyleIndex = -1;
    /**
     * 表格头的样式索引, -1表示未设置
     */
    protected int headerStyleIndex = -1;
    /**
     * 列宽，表头行高
     */
    public double width = -1D,
    /**
     * 表头行高
     */
    headerHeight = -1D;
    /**
     * 全局样式对象
     */
    public Styles styles;
    /**
     * 表头批注
     */
    public Comment headerComment, cellComment;
    /**
     * 表格体格式化，自定义格式化可以覆写方法{@link NumFmt#calcNumWidth(double, Font)}调整计算
     */
    public NumFmt numFmt;
    /**
     * 列索引，从0开始的数字，0对应Excel的{@code 'A'}列以此类推，-1表示未设置
     */
    public int colIndex = -1;
    /**
     * 多行表头中指向前一个{@code Column}
     *
     * @since 0.5.1
     */
    public Column prev;
    /**
     * 多行表头中指向后一个{@code Column}
     *
     * @since 0.5.1
     */
    public Column next;
    /**
     * 多行表头中指向最后一个{@code Column}
     *
     * @since 0.5.1
     */
    public Column tail;
    /**
     * 实际列索引，它与Excel列号对应，该值通过内部计算而来，请不要在外部修改
     */
    public int realColIndex;
    /**
     * 标志位集合，保存一些简单的标志位以节省空间，对应的位点说明如下
     *
     * <blockquote><pre>
     *  Bit  | Contents
     * ------+---------
     * 31, 1 | 自动换行 1位
     * 30, 2 | 自适应列宽 2位, 0: auto, 1: auto-size 2: fixed-size
     * 28, 1 | 忽略导出值 1位, 仅导出表头
     * 27, 1 | 隐藏列 1位
     * 26, 1 | 共享字符串 1位
     * 25, 2 | 列类型, 0: 默认导出为文本 1: 导出为图片 2: 超链接
     * 23, 2 | 垂直对齐
     * 21, 3 | 水平对齐
     * </pre></blockquote>
     */
    protected int option;
    /**
     * 图片效果，可以简单使用内置的{@link org.ttzero.excel.drawing.PresetPictureEffect} 28种效果
     */
    public Effect effect;
    /**
     * 单元格字体
     */
    public Font font;
    /**
     * 单元格填充
     */
    public Fill fill;
    /**
     * 单元格边框
     */
    public Border border;
    /**
     * 实例化列信息
     */
    public Column() { }

    /**
     * 实例化列信息
     *
     * @param name 列名，列名对应Excel表头
     */
    public Column(String name) {
        this.name = name;
    }

    /**
     * 实例化列信息
     *
     * @param name  列名，列名对应Excel表头
     * @param clazz 数据类型，影响单元格对齐默认字符串左对齐、数字右对齐、日期居中
     */
    public Column(String name, Class<?> clazz) {
        this(name, clazz, false);
    }

    /**
     * 实例化列信息
     *
     * @param name 列名，列名对应Excel表头
     * @param key  取值使用的关键字对应Java对象字段名或者Map的Key
     */
    public Column(String name, String key) {
        this(name, key, false);
    }

    /**
     * 实例化列信息
     *
     * @param name  列名，列名对应Excel表头
     * @param key   取值使用的关键字对应Java对象字段名、Map的Key或者SQL中select语句包含的字段
     * @param clazz 数据类型，影响单元格对齐默认字符串左对齐、数字右对齐、日期居中
     */
    public Column(String name, String key, Class<?> clazz) {
        this(name, key, false);
        this.clazz = clazz;
    }

    /**
     * 实例化列信息
     *
     * @param name  列名，列名对应Excel表头
     * @param clazz 数据类型，影响单元格对齐默认字符串左对齐、数字右对齐、日期居中
     * @param processor 输出转换器，动态转换状态值或枚举值为文本
     */
    public Column(String name, Class<?> clazz, ConversionProcessor processor) {
        this(name, clazz, processor, false);
    }

    /**
     * 实例化列信息
     *
     * @param name  列名，列名对应Excel表头
     * @param key   取值使用的关键字对应Java对象字段名、Map的Key或者SQL中select语句包含的字段
     * @param processor 输出转换器，动态转换状态值或枚举值为文本
     */
    public Column(String name, String key, ConversionProcessor processor) {
        this(name, key, processor, false);
    }

    /**
     * 实例化列信息
     *
     * @param name  列名，列名对应Excel表头
     * @param clazz 数据类型，影响单元格对齐默认字符串左对齐、数字右对齐、日期居中
     * @param share 是否将值放到字符串共享区
     */
    public Column(String name, Class<?> clazz, boolean share) {
        this.name = name;
        this.clazz = clazz;
        setShare(share);
    }

    /**
     * 实例化列信息
     *
     * @param name  列名，列名对应Excel表头
     * @param key   取值使用的关键字对应Java对象字段名、Map的Key或者SQL中select语句包含的字段
     * @param share 是否将值放到字符串共享区
     */
    public Column(String name, String key, boolean share) {
        this.name = name;
        this.key = key;
        setShare(share);
    }

    /**
     * 实例化列信息
     *
     * @param name  列名，列名对应Excel表头
     * @param clazz 数据类型，影响单元格对齐默认字符串左对齐、数字右对齐、日期居中
     * @param processor 输出转换器，动态转换状态值或枚举值为文本
     * @param share 是否将值放到字符串共享区
     */
    public Column(String name, Class<?> clazz, ConversionProcessor processor, boolean share) {
        this(name, clazz, share);
        this.processor = processor;
    }

    /**
     * 实例化列信息
     *
     * @param name  列名，列名对应Excel表头
     * @param key   取值使用的关键字对应Java对象字段名、Map的Key或者SQL中select语句包含的字段
     * @param clazz 数据类型，影响单元格对齐默认字符串左对齐、数字右对齐、日期居中
     * @param processor 输出转换器，动态转换状态值或枚举值为文本
     */
    public Column(String name, String key, Class<?> clazz, ConversionProcessor processor) {
        this(name, key, clazz);
        this.processor = processor;
    }

    /**
     * 实例化列信息
     *
     * @param name  列名，列名对应Excel表头
     * @param key   取值使用的关键字对应Java对象字段名、Map的Key或者SQL中select语句包含的字段
     * @param processor 输出转换器，动态转换状态值或枚举值为文本
     * @param share 是否将值放到字符串共享区
     */
    public Column(String name, String key, ConversionProcessor processor, boolean share) {
        this(name, key, share);
        this.processor = processor;
    }

    /**
     * 实例化列信息
     *
     * @param name  列名，列名对应Excel表头
     * @param clazz 数据类型，影响单元格对齐默认字符串左对齐、数字右对齐、日期居中
     * @param cellStyle 样式值，样式值由背景，边框，字体等进行“或”运算而来
     */
    public Column(String name, Class<?> clazz, int cellStyle) {
        this(name, clazz, cellStyle, true);
    }

    /**
     * 实例化列信息
     *
     * @param name  列名，列名对应Excel表头
     * @param key   取值使用的关键字对应Java对象字段名、Map的Key或者SQL中select语句包含的字段
     * @param cellStyle 样式值，样式值由背景，边框，字体等进行“或”运算而来
     */
    public Column(String name, String key, int cellStyle) {
        this(name, key, cellStyle, true);
    }

    /**
     * 实例化列信息
     *
     * @param name  列名，列名对应Excel表头
     * @param clazz 数据类型，影响单元格对齐默认字符串左对齐、数字右对齐、日期居中
     * @param cellStyle 样式值，样式值由背景，边框，字体等进行“或”运算而来
     * @param share 是否将值放到字符串共享区
     */
    public Column(String name, Class<?> clazz, int cellStyle, boolean share) {
        this(name, clazz, share);
        this.cellStyle = cellStyle;
    }

    /**
     * 实例化列信息
     *
     * @param name  列名，列名对应Excel表头
     * @param key   取值使用的关键字对应Java对象字段名、Map的Key或者SQL中select语句包含的字段
     * @param cellStyle 样式值，样式值由背景，边框，字体等进行“或”运算而来
     * @param share 是否将值放到字符串共享区
     */
    public Column(String name, String key, int cellStyle, boolean share) {
        this(name, key, share);
        this.cellStyle = cellStyle;
    }

    /**
     * 通过已有列实例化列信息
     *
     * @param other 其它列
     */
    public Column(Column other) {
        from(other);
        if (other.next != null) addSubColumn(new Column(other.next));
    }

    /**
     * 复制列信息
     *
     * @param other 其它列
     * @return 当前列 
     */
    public Column from(Column other) {
        this.key = other.key;
        this.name = other.name;
        this.clazz = other.clazz;
        this.processor = other.processor;
        this.converter = other.converter;
        this.styleProcessor = other.styleProcessor;
        this.width = other.width;
        this.headerHeight = other.headerHeight;
        this.styles = other.styles;
        this.headerComment = other.headerComment;
        this.cellComment = other.cellComment;
        this.numFmt = other.numFmt;
        this.font = other.font;
        this.border = other.border;
        this.fill = other.fill;
        this.colIndex = other.colIndex;
        this.option = other.option;
        this.realColIndex = other.realColIndex;
        if (other.cellStyle != null) setCellStyle(other.cellStyle);
        if (other.headerStyle != null) setHeaderStyle(other.headerStyle);
        int i;
        if ((i = other.getHeaderStyleIndex()) > 0) this.headerStyleIndex = i;
        if ((i = other.getCellStyleIndex()) > 0) this.cellStyleIndex = i;
        this.effect = other.effect;

        return this;
    }

    /**
     * 设置列宽，当该列同时标记为“自适应列宽”时最终的列宽取两者中的较小值，当列宽设置为{@code 0}时效果与隐藏相同
     *
     * @param width 列宽，必须大于等于0
     * @return 当前列
     */
    public Column setWidth(double width) {
        if (width < 0) {
            throw new ExcelWriteException("Width " + width + " less than 0.");
        }
        this.width = width;
        return this;
    }

    /**
     * 设置行高，最终的行高取所有列最大值，当行高设置为{@code 0}时效果与隐藏相同
     *
     * @param headerHeight 行高，必须大于等于{@code 0}
     * @return 当前列
     */
    public Column setHeaderHeight(double headerHeight) {
        if (headerHeight < 0) {
            throw new ExcelWriteException("Height " + headerHeight + " less than 0.");
        }
        this.headerHeight = headerHeight;
        return this;
    }

    /**
     * 获取是否将字符串放入共享区
     *
     * @return true: 共享，false：内嵌
     */
    public boolean isShare() {
        return (option >> 5 & 1) == 1;
    }

    /**
     * 获取列名，列名对应Excel表头
     *
     * @return 列名
     */
    public String getName() {
        return name;
    }

    /**
     * 设置表头列名
     *
     * @param name 表头列名
     * @return 当前列
     */
    public Column setName(String name) {
        this.name = name;
        return this;
    }

    /**
     * 获取列数据类型
     *
     * @return 数据行的数据类型
     */
    public Class<?> getClazz() {
        return clazz;
    }

    /**
     * 设置列数据类型，数据影响单元格对齐，默认字符串左对齐、数字右对齐、日期居中
     *
     * @param clazz 列数据类型
     * @return 当前列
     */
    public Column setClazz(Class<?> clazz) {
        this.clazz = clazz;
        return this;
    }

    /**
     * 设置输出转换器，通常用于动态转换状态值或枚举值为文本
     *
     * @param processor 输出转换器
     * @return 当前列
     */
    public Column setProcessor(ConversionProcessor processor) {
        this.processor = processor;
        return this;
    }

    /**
     * 设置动态样式转换器，通常用于高亮显示单元格起提醒作用
     *
     * @param styleProcessor 样式转换器
     * @return 当前列
     */
    public Column setStyleProcessor(StyleProcessor styleProcessor) {
        if (styleProcessor != null && !StyleProcessor.None.class.isAssignableFrom(styleProcessor.getClass())) {
            this.styleProcessor = styleProcessor;
        }
        return this;
    }

    /**
     * 获取输出转换器，优先返回{@code ConversionProcessor}，其次是{@code Converter}
     *
     * @return 值转换器
     */
    public ConversionProcessor getConversion() {
        return processor != null ? processor : converter;
    }

    /**
     * 设置转换器，导出的时候将状态值或枚举值转为文本，导入的时候将文本转为状态或枚举值
     *
     * @param converter 值转换器
     * @return 当前列
     */
    public Column setConverter(Converter<?> converter) {
        if (converter != null && !Converter.None.class.isAssignableFrom(converter.getClass())) {
            this.converter = converter;
        }
        return this;
    }

    /**
     * 获取列宽，这里仅返回通过{@link #setWidth}设置的列宽不包含自适应列宽，自适应列宽通常是在导出的时候动态计算
     *
     * @return 当前列宽，未设置时返回{@code -1}
     */
    public double getWidth() {
        return width;
    }

    /**
     * 设置单元格样式值，样式值由背景，边框，字体等进行“或”运算而来
     *
     * @param cellStyle 样式值
     * @return 当前列
     */
    public Column setCellStyle(int cellStyle) {
        this.cellStyle = cellStyle;
        if (styles != null) this.cellStyleIndex = styles.of(cellStyle);
        return this;
    }

    /**
     * 设置表头单元格样式值，样式值由背景，边框，字体等进行“或”运算而来
     *
     * @param headerStyle 样式值
     * @return 当前列
     */
    public Column setHeaderStyle(int headerStyle) {
        this.headerStyle = headerStyle;
        if (styles != null) this.headerStyleIndex = styles.of(headerStyle);
        return this;
    }

    /**
     * 设置列下标，下标从{@code 0}开始对应Excel的{@code A}列，这里设置的下标是绝对位置，
     * 如果表头下标不连续那么导出的时候列也是不连续的
     *
     * @param colIndex 从0开始的列号
     * @return 当前列
     */
    public Column setColIndex(int colIndex) {
        this.colIndex = colIndex;
        return this;
    }

    /**
     * 获取单元格样式索引，不包含动态样式
     *
     * @return 单元格样式索引，未设置样式时返回{@code -1}
     */
    public int getCellStyleIndex() {
        return cellStyleIndex >= 0 ? cellStyleIndex : (cellStyleIndex = styles != null && cellStyle != null ? styles.of(cellStyle) : -1);
    }

    /**
     * 获取表头单元格样式索引，不包含动态样式
     *
     * @return 表头单元格样式索引，未设置样式时返回{@code -1}
     */
    public int getHeaderStyleIndex() {
        return headerStyleIndex >= 0 ? headerStyleIndex : (headerStyleIndex = styles != null && headerStyle != null ? styles.of(headerStyle) : -1);
    }

    /**
     * 获取默认水平对齐，水平对齐可用值包含在{@link Horizontals}定义中，
     * 默认的时间居中，数字居右其余居左
     *
     * @return 水平对齐
     */
    protected int defaultHorizontal() {
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
     * 设置当前列统一“字体”样式
     *
     * @param font 字体
     * @return 当前列
     */
    public Column setFont(Font font) {
        this.font = font;
        return this;
    }

    /**
     * 设置当前列统一“填充”样式
     *
     * @param fill 填充
     * @return 当前列
     */
    public Column setFill(Fill fill) {
        this.fill = fill;
        return this;
    }

    /**
     * 设置当前列统一“边框”样式
     *
     * @param border 边框
     * @return 当前列
     */
    public Column setBorder(Border border) {
        this.border = border;
        return this;
    }

    /**
     * 设置当前列统一“垂直对齐”样式
     *
     * @param vertical 垂直对齐，参考值{@link Verticals}
     * @return 当前列
     */
    public Column setVertical(int vertical) {
        option = (option & ~(3 << 8)) | (((vertical >> Styles.INDEX_VERTICAL) & 3) << 8);
        return this;
    }

    /**
     * 设置当前列统一“水平对齐”样式
     *
     * @param horizontal 水平对齐,参考值{@link Horizontals}
     * @return 当前列
     */
    public Column setHorizontal(int horizontal) {
        option = (option & ~(7 << 10)) | (((horizontal >> Styles.INDEX_HORIZONTAL) & 7) << 10);
        return this;
    }

    /**
     * 设置共享字符串标记，当此标记为{@code true}时单元格的字符串将独立保存在共享区
     *
     * @param share true: 共享, false: 内嵌
     * @return 当前列
     */
    public Column setShare(boolean share) {
        if (share) this.option |= 1 << 5;
        else this.option &= ~(1 << 5);
        return this;
    }

    /**
     * 设置当前列统一“格式化”样式
     *
     * @param code 格式化串
     * @return 当前列
     */
    public Column setNumFmt(String code) {
        this.numFmt = new NumFmt(code);
        return this;
    }

    /**
     * 设置当前列统一“格式化”样式
     *
     * @param numFmt 格式化{@link NumFmt}
     * @return 当前列
     */
    public Column setNumFmt(NumFmt numFmt) {
        this.numFmt = numFmt;
        return this;
    }

    /**
     * 获取当前列统一“格式化”，仅返回通过{@link #setNumFmt}设置的统一值，不包含动态样式
     *
     * @return 格式化 {@link NumFmt}
     */
    public NumFmt getNumFmt() {
        return numFmt != null ? numFmt : (numFmt = styles.getNumFmt(cellStyle));
    }

    /**
     * 获取单元格默认样式，通常该方法仅初始化时被调用一次
     *
     * @param clazz 列数据类型
     * @return 样式值
     */
    protected int getCellStyle(Class<?> clazz) {
        int style;
        if (isString(clazz)) {
            style = Styles.defaultStringBorderStyle();
        } else if (isDateTime(clazz) || isDate(clazz) || isLocalDateTime(clazz)) {
            if (numFmt == null) numFmt = DATETIME_FORMAT;
            style = (1 << Styles.INDEX_FONT) | (1 << INDEX_BORDER) | Horizontals.CENTER;
        } else if (isBool(clazz) || isChar(clazz)) {
            style = Styles.clearHorizontal(Styles.defaultStringBorderStyle()) | Horizontals.CENTER;
        } else if (isInt(clazz) || isLong(clazz)) {
            style = Styles.defaultIntBorderStyle();
        } else if (isFloat(clazz) || isDouble(clazz) || isBigDecimal(clazz)) {
            style = Styles.defaultDoubleBorderStyle();
        } else if (isLocalDate(clazz)) {
            if (numFmt == null) numFmt = DATE_FORMAT;
            style = (1 << Styles.INDEX_FONT) | (1 << INDEX_BORDER) | Horizontals.CENTER;
        } else if (isTime(clazz) || isLocalTime(clazz)) {
            if (numFmt == null) numFmt = TIME_FORMAT;
            style = (1 << Styles.INDEX_FONT) | (1 << INDEX_BORDER) | Horizontals.CENTER;
        } else {
            style = (1 << Styles.INDEX_FONT) | (1 << INDEX_BORDER); // Auto-style
        }

        return style;
    }

    /**
     * 获取单元格样式值，导出过程中通常只计算一次通用样式，计算完后值将保存在{@code cellStyle}中，
     * 默认样式做为底色外部设置的“字体”，“填充”等样式将覆盖默认样式，
     *
     * @return 样式值
     */
    public int getCellStyle() {
        if (cellStyle != null) return cellStyle;
        // 获取默认样式
        int style = getCellStyle(clazz);

        // 重置"字体"
        if (font != null) style = styles.modifyFont(style, font);
        // 重置“格式化”
        if (numFmt != null) style = styles.modifyNumFmt(style, numFmt);
        // 重置“边框”
        if (border != null) style = styles.modifyBorder(style, border);
        // 重置“填充”
        if (fill != null) style = styles.modifyFill(style, fill);
        // 重置“垂直对齐”
        int v = ((option >>> 8) & 3) << Styles.INDEX_VERTICAL;
        if (v > 0) style = styles.modifyVertical(style, v);
        // 重置“水平对齐”
        int h = ((option >>> 10) & 7) << Styles.INDEX_HORIZONTAL;
        if (h > 0) style = styles.modifyHorizontal(style, h);
        // 重置“自动折行”
        style |= (option & 1);

        // 保存样式
        setCellStyle(style);
        return style;
    }

    /**
     * 是否忽略数据
     *
     * @return true: 忽略数据只输出表头
     */
    public boolean isIgnoreValue() {
        return (option >> 3 & 1) == 1;
    }

    /**
     * 忽略{@code Body}的数据只输出表头
     *
     * @return 当前列
     */
    public Column ignoreValue() {
        this.option |= 1 << 3;
        return this;
    }

    /**
     * 设置“自动折行”
     *
     * <p>折行触发条件：一是当长度超过列宽时折行，二是包含回车符时折行</p>
     *
     * @param wrapText 自动折行
     * @return 当前列
     */
    public Column setWrapText(boolean wrapText) {
        if (wrapText) this.option |= 1;
        else this.option = option >>> 1 << 1;
        return this;
    }

    /**
     * 设置表头批注
     *
     * @param headerComment 批注{@link Comment}
     * @return 当前列
     */
    public Column setHeaderComment(Comment headerComment) {
        this.headerComment = headerComment;
        return this;
    }

    /**
     * 在尾部添加表头
     *
     * @param column 表头
     * @return 当前列
     */
    public Column addSubColumn(Column column) {
        if (this == column) return this;
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
     * 获取表头行数，单表头返回1，多表头时返回表头行数
     *
     * @return 表头行数
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
     * 多表头时将链表转为数组
     *
     * @return 表头
     */
    public Column[] toArray() {
        int len = subColumnSize();
        Column[] dist = new Column[len];
        if (len < 1) return dist;
        Column e = this;
        for (int i = 0; i < len; i++) {
            dist[i] = e;
            e = e.next;
        }
        return dist;
    }

    /**
     * 获取列在Excel中的实际位置（从{@code A}开始）
     *
     * @return 列在Excel中的实际位置
     */
    public int getRealColIndex() {
        return realColIndex;
    }

    /**
     * 判断当前列是否“隐藏”
     *
     * @return true: 隐藏
     */
    public boolean isHide() {
        return (option >> 4 & 1) == 1;
    }

    /**
     * 隐藏当前列
     *
     * @return 当前列
     */
    public Column hide() {
        this.option |= 1 << 4;
        return this;
    }

    /**
     * 标记当前列可见（默认可见）
     *
     * @return 当前列
     */
    public Column show() {
        this.option &= ~(1 << 4);
        return this;
    }

    /**
     * 获取尾列，Excel从上到下记为首-尾列，尾列为最接近表格体{@code Body}的列
     *
     * @return 尾列
     */
    public Column getTail() {
        return tail != null ? tail : this;
    }

    /**
     * 标记“自适应”列宽，导出时根据单元格内容动态计算列宽
     *
     * @return 当前列
     */
    public Column autoSize() {
        this.option |= 1 << 1;
        return this;
    }

    /**
     * 设置固定列宽
     *
     * @param width 列宽，必须大于等于0
     * @return 当前列
     */
    public Column fixedSize(double width) {
        this.option |= 1 << 2;
        this.width = width;
        return this;
    }

    /**
     * 获取列宽属性
     *
     * @return 0: 未设置 1: 自适应列宽 2: 固定列宽
     */
    public int getAutoSize() {
        return option >> 1 & 3;
    }

    /**
     * 指定当前列以“值”类型导出
     *
     * @return 当前列
     */
    public Column writeAsDefault() {
        this.option &= ~(3 << 6);
        return this;
    }

    /**
     * 指定当前列以“媒体”类型导出
     *
     * @return 当前列
     */
    public Column writeAsMedia() {
        this.option = this.option & ~(3 << 6) | (1 << 6);
        return this;
    }

    /**
     * 获取列属性
     *
     * @return 0: 默认 1: 媒体（图片） 2: 超链接
     */
    public int getColumnType() {
        return (this.option >> 6) & 3;
    }

    /**
     * 设置当前列全局图片效果，只有当{@code columnType}为{@code Media}时生效
     *
     * @param effect 图片效果{@link Effect}
     * @return 当前列
     */
    public Column setEffect(Effect effect) {
        this.effect = effect;
        return this;
    }

    /**
     * 获取当前列设置的图片效果
     *
     * @return 图片效果 {@link Effect}
     */
    public Effect getEffect() {
        return effect;
    }

    /**
     * 指定当前列以“超链接”类型导出
     *
     * @return 当前列
     */
    public Column writeAsHyperlink() {
        this.option = this.option & ~(3 << 6) | (2 << 6);
        return this;
    }
}

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
     * 31. 1 | 自动换行 1位
     * 30. 2 | 自适应列宽 2位, 0: auto, 1: auto-size 2: fixed-size
     * 28. 1 | 忽略导出值 1位, 仅导出表头
     * 27. 1 | 隐藏列 1位
     * 26. 1 | 共享字符串 1位
     * 25. 2 | 列类型, 0: 默认导出为文本 1: 导出为图片
     * </pre></blockquote>
     */
    protected int option;
    /**
     * 图片效果
     */
    public Effect effect;
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
        this(name, clazz, false);
    }

    /**
     * Constructor Column
     *
     * @param name the column name
     * @param key  field
     */
    public Column(String name, String key) {
        this(name, key, false);
    }

    /**
     * Constructor Column
     *
     * @param name  the column name
     * @param key   field
     * @param clazz the cell type
     */
    public Column(String name, String key, Class<?> clazz) {
        this(name, key, false);
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
        this(name, clazz, processor, false);
    }

    /**
     * Constructor Column
     *
     * @param name      the column name
     * @param key       field
     * @param processor The int value conversion
     */
    public Column(String name, String key, ConversionProcessor processor) {
        this(name, key, processor, false);
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
        setShare(share);
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
        setShare(share);
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
     * Create column from exists column
     *
     * @param other the other column
     */
    public Column(Column other) {
        from(other);
        if (other.next != null) addSubColumn(new Column(other.next));
    }

    /**
     * Copy properties from other column
     *
     * @param other the other column
     * @return current
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
     * Setting the cell's width
     *
     * @param width the width value
     * @return the {@link Column}
     */
    public Column setWidth(double width) {
        if (width < 0) {
            throw new ExcelWriteException("Width " + width + " less than 0.");
        }
        this.width = width;
        return this;
    }

    /**
     * Setting the cell's height, The row height is equal to the maximum cell height
     *
     * @param headerHeight the header height value
     * @return the {@link Column}
     */
    public Column setHeaderHeight(double headerHeight) {
        if (headerHeight < 0) {
            throw new ExcelWriteException("Height " + headerHeight + " less than 0.");
        }
        this.headerHeight = headerHeight;
        return this;
    }

    /**
     * Setting the cell is shared
     *
     * @return true:shared false:inline string
     */
    public boolean isShare() {
        return (option >> 5 & 1) == 1;
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
     * 设置转换器
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
        if (share) this.option |= 1 << 5;
        else this.option &= ~(1 << 5);
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
            style = Styles.defaultStringBorderStyle();
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
            style = styles.modifyNumFmt(style, numFmt);
        }

        return style | (option & 1);
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
        return (option >> 3 & 1) == 1;
    }

    /**
     * Ignore value
     *
     * @return the {@link Column} self
     */
    public Column ignoreValue() {
        this.option |= 1 << 3;
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
        if (wrapText) this.option |= 1;
        else this.option = option >>> 1 << 1;
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
     * @return header columns
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
     *
     * @return real col-index(one base)
     */
    public int getRealColIndex() {
        return realColIndex;
    }

    /**
     * Returns hide flag
     *
     * @return true: hidden otherwise show
     */
    public boolean isHide() {
        return (option >> 4 & 1) == 1;
    }

    /**
     * Hidden current column
     *
     * @return current {@link Column}
     */
    public Column hide() {
        this.option |= 1 << 4;
        return this;
    }

    /**
     * Show current column
     *
     * @return current {@link Column}
     */
    public Column show() {
        this.option &= ~(1 << 4);
        return this;
    }

    /**
     * Returns the last column
     *
     * @return the last column (real column)
     */
    public Column getTail() {
        return tail != null ? tail : this;
    }

    /**
     * Setting auto resize cell's width
     *
     * @return current {@link Column}
     */
    public Column autoSize() {
        this.option |= 1 << 1;
        return this;
    }

    /**
     * Setting fix column width
     *
     * @return current {@link Column}
     */
    public Column fixedSize() {
        this.option |= 1 << 2;
        return this;
    }

    /**
     * Setting fixed column width
     *
     * @param width the column width
     * @return current {@link Column}
     */
    public Column fixedSize(double width) {
        this.option |= 1 << 2;
        this.width = width;
        return this;
    }

    /**
     * Returns the re-size setting
     *
     * @return 0: not setting 1: auto-size 2:fixed-size
     */
    public int getAutoSize() {
        return option >> 1 & 3;
    }

    /**
     * Write cell value as default
     *
     * @return current {@link Column}
     */
    public Column writeAsDefault() {
        this.option &= ~(3 << 6);
        return this;
    }

    /**
     * Specify the cell type as media (drawing,chart eq.)
     *
     * @return current {@link Column}
     */
    public Column writeAsMedia() {
        this.option = this.option & ~(3 << 6) | (1 << 6);
        return this;
    }

    /**
     * Returns the column type
     *
     * @return 0: default 1: picture
     */
    public int getColumnType() {
        return (this.option >> 6) & 3;
    }

    /**
     * Specify the image effect, which only takes effect when the column write as Media
     *
     * @param effect {@link Effect}
     * @return current {@link Column}
     */
    public Column setEffect(Effect effect) {
        this.effect = effect;
        return this;
    }

    /**
     * Returns {@code Effect}
     *
     * @return {@link Effect}
     */
    public Effect getEffect() {
        return effect;
    }
}

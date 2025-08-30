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


import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.ttzero.excel.entity.e7.XMLWorksheetWriter;
import org.ttzero.excel.entity.style.Border;
import org.ttzero.excel.entity.style.BorderStyle;
import org.ttzero.excel.entity.style.Fill;
import org.ttzero.excel.entity.style.Font;
import org.ttzero.excel.entity.style.Horizontals;
import org.ttzero.excel.entity.style.NumFmt;
import org.ttzero.excel.entity.style.PatternType;
import org.ttzero.excel.entity.style.Styles;
import org.ttzero.excel.entity.style.Verticals;
import org.ttzero.excel.manager.Const;
import org.ttzero.excel.manager.RelManager;
import org.ttzero.excel.reader.Cell;
import org.ttzero.excel.reader.Dimension;
import org.ttzero.excel.util.FileUtil;

import java.awt.Color;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.ObjectInputStream;
import java.io.ObjectOutputStream;
import java.lang.reflect.InvocationTargetException;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.function.BiConsumer;

import static org.ttzero.excel.manager.Const.ROW_BLOCK_SIZE;
import static org.ttzero.excel.reader.ExcelReader.coordinateToLong;
import static org.ttzero.excel.util.ExtBufferedWriter.getChars;
import static org.ttzero.excel.util.ExtBufferedWriter.stringSize;
import static org.ttzero.excel.util.StringUtil.isEmpty;
import static org.ttzero.excel.util.StringUtil.isNotEmpty;

/**
 * 工作表Worksheet是Excel最重要的组件，在Excel看见的所有内容都是由Worksheet工作表呈现。
 * 本工具将工作表{@code Sheet}及其子类视为数据源，它本身除了收集数据外并不输出任何格式的文件，
 * 它必须与输出协议搭配使用才会输出相应格式的文件。如与{@link XMLWorksheetWriter}输出协议搭配时，
 * 每个{@code Sheet}对应一个Excel工作表和{@code sheet.xml}文件。
 *
 * <p>工作表对应的输出协议为{@link IWorksheetWriter}，它会循环调用{@link #nextBlock}
 * 方法获取数据并写入磁盘直到{@link RowBlock#isEOF}返回EOF标记为止，整个过程只有一个
 * RowBlock行块常驻内存，一个{@code RowBlock}行块默认包含32个{@code Row}行，这样可以保证
 * 较小的内存开销。</p>
 *
 * <p>当前支持的数据源有{@link ListSheet},{@link ListMapSheet},{@link TemplateSheet},{@link StatementSheet}
 * 和{@link ResultSetSheet}5种，前三种较为常用，后两种可实现将数据库查询结果直接导出到Excel
 * 省掉转Java实体的中间环节。继承Sheet并实现抽象方法{@link Sheet#resetBlockData}可以扩展新的数据源，
 * 你需要在该方法中获取数据并使用{@link ICellValueAndStyle}转换器将数据转换为输出协议允许的结构。</p>
 *
 * <p>{@code ListSheet}及其子类支持超大数据导出，理论上导出数据行无上限，当超过Worksheet行数限制时
 * 将会在下一个位置插入一个与当前{@code Sheet}一样的工作表，然后将剩余数据写到新插入的工作表中，
 * 分页是自动触发的无需额外设置。超大数据导出时建议使用{@link ListSheet#more}方法分批查询数据，
 * 返回空数组或{@code null}时结束。</p>
 *
 * <p>每个Worksheet都可以设置一个{@link #onProgress}窗口来观察导出进度，通过此窗口可以记录每个
 * RowBlock导出时间然后预估整体的导出时间。</p>
 *
 * <p>能被直接导出的类型包含Java定义的简单类型、包装类型以及时间类型，除此之外的其它类型均会调用{@code toString}
 * 后直接输出，所以自定义类型或枚举可以覆写{@code toString}方法转为可读的字符串，或者使用
 * {@link org.ttzero.excel.processor.ConversionProcessor}动态转换将不可读的状态值{@code 1,2,3}转换为
 * 可读的文字”申请中“，“二申中”，“通过”等，对于未知的类型还可以实现{@link ICellValueAndStyle}转换器并覆写
 * {@link ICellValueAndStyle#unknownType}方法进行转换</p>
 *
 * <p>关于扩展属性：随着功能越加越多，在{@code Sheet}中定义的属性也越来越多，这样无限添加可不是个好主意，
 * 所以在{@code v0.5.0}引入了一个{@code Map}类型的扩展属性{@link #extProp}，通过{@link #putExtProp}
 * 和{@link #getExtPropValue}方法添加和读取扩展属性，一般情况下在数据源中添加属性，在输出协议中读取属性，
 * 像合并单元格、冻结首行都是通过扩展参数实现。</p>
 *
 * @author guanquan.wang on 2017/9/26.
 * @see ListSheet
 * @see ListMapSheet
 * @see TemplateSheet
 * @see ResultSetSheet
 * @see StatementSheet
 * @see CSVSheet
 * @see SimpleSheet
 */
public abstract class Sheet implements Cloneable, Storable {
    /**
     * LOGGER
     */
    protected final Logger LOGGER = LoggerFactory.getLogger(getClass());
    /**
     * 工作薄
     */
    protected Workbook workbook;
    /**
     * 工作表名称
     */
    protected String name;
    /**
     * 表头
     */
    protected Column[] columns;
    /**
     * 水印
     */
    protected WaterMark waterMark;
    /**
     * 关系管理器
     */
    protected RelManager relManager;
    /**
     * 工作表ID，与当前工作表在工作薄中的下标一致
     */
    protected int id;
    /**
     * 表头批注
     */
    protected Comments comments;
    /**
     * 自适应列宽标记，优先级从小到大为 0: 未设置 1: 自适应列宽 2: 固定宽度
     */
    protected int autoSize;
    /**
     * 默认列宽
     */
    protected double width = 20D;
    /**
     * 统计已写入数据行数，不包含表头
     */
    protected int rows;
    /**
     * 标记是否“隐藏”
     */
    protected boolean hidden;
    /**
     * 兜底的表头样式索引，优先级低Column独立设置的样式
     */
    protected int headStyleIndex = -1;
    /**
     * 统一的表头样式，优先级低Column独立设置的样式
     */
    protected int headStyle;
    /**
     * 斑马线样式索引
     */
    protected int zebraFillStyle = -1;
    /**
     * 斑马线填充样式，斑马线从表头以下的第2行开始每隔一行进行一次填充
     */
    protected Fill zebraFill;
    /**
     * 标记是否为自动分页的“复制”工作表
     */
    protected boolean copySheet;
    /**
     * 记录自动分页的“复制”工作表数量
     */
    protected int copyCount;
    /**
     * 行块，它由连续的一组默认大小为{@code 32}个{@code Row}组成的迭代器，该对象是内存共享的，
     * 可以通过覆写{@link #getRowBlockSize()}指定其它大小，一般不建议修改。
     */
    protected RowBlock rowBlock;
    /**
     * 工作表输出协议
     */
    protected IWorksheetWriter sheetWriter;
    /**
     * 标记表头是否已采集，默认情况下会进行收集-排序-多表头合并等过程后状态才为"ready"
     */
    protected boolean headerReady;
    /**
     * 标记是否需要关闭，自动分页情况下最后一个worksheet页需要关闭资源，因为所有的数据都是从最原始
     * 的工作表获取，所以只有写完数据之后才能关闭。
     */
    protected boolean shouldClose = true;
    /**
     * 转换器，将外部数据转换为Worksheet输出协议需要的数据类型并设置单元格样式
     */
    protected ICellValueAndStyle cellValueAndStyle;
    /**
     * 忽略表头 -1 未设置, 0 输出表头, 1 忽略表头
     */
    protected int nonHeader = -1;
    /**
     * 工作表body行数上限，它记录的是输出协议行上限-表头行数，例如xls格式最多{@code 65535}行，
     * 导出的表头为1行，那{@code rowLimit = 65534}
     */
    private int rowLimit;
    /**
     * 扩展属性
     */
    protected Map<String, Object> extProp = new HashMap<>();
    /**
     * 扩展参数的位标志。如果存在扩展参数则相应的位为1，低16位由系统占用
     */
    protected int extPropMark;
    /**
     * 是否显示"网格线"，默认显示
     */
    protected Boolean showGridLines;
    /**
     * 指定表头行高
     */
    protected double headerRowHeight = 20.5D;
    /**
     * 指定数据行高
     */
    protected Double rowHeight;
    /**
     * 指定起始行列，高48位保存Row，低16位保存Col(zero base)
     */
    protected long startCoordinate;
    /**
     * 导出进度窗口，默认情况下RowBlock每刷新一次就会更新一次进度，也就是每32行通知一次
     */
    protected BiConsumer<Sheet, Integer> progressConsumer;
    /**
     * 获取工作表ID，与当前工作表在工作薄中的下标一致，一般与其它资源关联使用
     *
     * @return 工作表ID
     */
    public int getId() {
        return id;
    }

    /**
     * 设置工作表ID，请不要在外部随意修改，否则打开文件异常
     *
     * @param id 工作表ID
     * @return 当前工作表
     */
    public Sheet setId(int id) {
        this.id = id;
        return this;
    }

    /**
     * 设置输出协议，必须与对应的{@link IWorkbookWriter}工作薄输出协议一起使用
     *
     * @param sheetWriter 工作表输出协议{@link IWorksheetWriter}
     * @return 当前工作表
     */
    public Sheet setSheetWriter(IWorksheetWriter sheetWriter) {
        this.sheetWriter = sheetWriter;
        this.sheetWriter.setWorksheet(this);
        return this;
    }

    /**
     * 获取工作表输出协议{@link IWorksheetWriter}
     *
     * @return 工作表输出协议
     */
    public IWorksheetWriter getSheetWriter() {
        return sheetWriter;
    }

    /**
     * 设置数据转换器，用于将Java对象转为各工作表输出协议可接受的数据结构，一般会将每个单元格的值输出为{@link Cell}对象，
     * 它与{@code Sheet}数据源中的Java类型完全分离，使得下游的输出协议有统一输入源。
     *
     * <p>除了数据转换外，该转换器还兼具采集样式，采集样式时会先从表头{@link Column#getCellStyle}中获取初始样式，
     * 如果该列有动态样式则会将该初始样式做为入参传入{@link org.ttzero.excel.processor.StyleProcessor}
     * 动态样式处理器以制定动态样式，如果设置有工作表级的样式处理器则会将动态样式的结果做为入参继续调工作表级
     * 样式处理器制定最终的样式</p>
     *
     * @param cellValueAndStyle 数据转换器
     * @return 当前工作表
     */
    public Sheet setCellValueAndStyle(ICellValueAndStyle cellValueAndStyle) {
        this.cellValueAndStyle = cellValueAndStyle;
        return this;
    }

    /**
     * 获取数据转换器
     *
     * @return 数据转换器 {@link ICellValueAndStyle}
     */
    public ICellValueAndStyle getCellValueAndStyle() {
        return cellValueAndStyle;
    }

    /**
     * 实例化工作表，未指定工作表名称时默认以{@code 'Sheet'+id}命名
     */
    public Sheet() {
        relManager = new RelManager();
    }

    /**
     * 实例化工作表并指定工作表名称
     *
     * @param name 工作表名称
     */
    public Sheet(String name) {
        this.name = name;
        relManager = new RelManager();
    }

    /**
     * 实例化工作表并指定表头信息
     *
     * @param columns 表头信息
     */
    public Sheet(final Column... columns) {
        this.columns = columns;
        relManager = new RelManager();
    }

    /**
     * 实例化工作表并指定工作表名称和表头信息
     *
     * @param name    工作表名称
     * @param columns 表头信息
     */
    public Sheet(String name, final Column... columns) {
        this(name, null, columns);
    }

    /**
     * 实例化工作表并指定工作表名称，水印和表头信息
     *
     * @param name      工作表名称
     * @param waterMark 水印
     * @param columns   表头信息
     */
    public Sheet(String name, WaterMark waterMark, final Column... columns) {
        this.name = name;
        this.columns = columns;
        this.waterMark = waterMark;
        relManager = new RelManager();
    }

    /**
     * 获取当前工作表对应的工作薄
     *
     * @return 当前工作表对应的 {@link Workbook}
     */
    public Workbook getWorkbook() {
        return workbook;
    }

    /**
     * 设置工作薄，一般在调用{@link Workbook#addSheet}时设置工作薄，{@code Workbook}包含
     * 样式、共享字符区、资源类型等全局配置，为了方便读取所以每个worksheet均包含Workbook句柄
     *
     * @param workbook 工作薄{@link Workbook}
     * @return 当前工作表
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
     * 获取默认列宽，如果未在Column上特殊指定宽度时该宽度将应用于每一列
     *
     * @return 默认列宽20
     */
    public double getDefaultWidth() {
        return width;
    }

    /**
     * 标记当前工作表自适应列宽，此优先级低于使用{@link Column#setWidth}指定列宽
     *
     * @return 当前工作表
     */
    public Sheet autoSize() {
        this.autoSize = 1;
        return this;
    }

    /**
     * 获取工作表的全局自适应列宽标记
     *
     * @return 1: 自适应列宽 2: 固定宽度
     */
    public int getAutoSize() {
        return autoSize;
    }

    /**
     * 是否全局自适应列宽，此值使用{@code autoSize()==1}判断
     *
     * @return true：自适应列宽
     */
    public boolean isAutoSize() {
        return autoSize == 1;
    }

    /**
     * 设置当前工作表使用固定列宽，将默认使用{@link #getDefaultWidth()}返回的宽度
     *
     * @return 当前工作表
     */
    public Sheet fixedSize() {
        this.autoSize = 2;
        return this;
    }

    /**
     * 设置当前工作表使用固定列宽并指定宽度，此方法会对入参进行重算，当宽度为'零'时效果相当于隐藏该列
     *
     * @param width 列宽
     * @return 当前工作表
     */
    public Sheet fixedSize(double width) {
        if (width < 0.0D) {
            LOGGER.warn("Negative number {}", width);
            width = 0.0D;
        }
        else if (width > Const.Limit.COLUMN_WIDTH) {
            LOGGER.warn("Maximum width is {}, current is {}", Const.Limit.COLUMN_WIDTH, width);
            width = Const.Limit.COLUMN_WIDTH;
        }
        this.autoSize = 2;
        this.width = width;
        if (headerReady) {
            for (org.ttzero.excel.entity.Column hc : columns) {
                hc.fixedSize(width);
            }
        }
        return this;
    }

    /**
     * 设置斑马线填充样式，为了不影响正常阅读建议使用浅色，默认无斑马线
     *
     * @param fill 斑马线填充 {@link Fill}
     * @return 当前工作表
     */
    public Sheet setZebraLine(Fill fill) {
        this.zebraFill = fill;
        return this;
    }

    /**
     * 取消斑马线，如果在工作薄Workbook设置了全局斑马线可使用此方法取消当前工作表Worksheet的斑马线
     *
     * @return 当前工作表
     */
    public Sheet cancelZebraLine() {
        this.zebraFill = null;
        this.zebraFillStyle = 0;
        return this;
    }

    /**
     * 获取当前工作表的斑马线填充样式，如果当前工作表未设置则从全局工作薄中获取
     *
     * @return 斑马线填充 {@link Fill}
     */
    public Fill getZebraFill() {
        return zebraFill != null ? zebraFill : workbook.getZebraFill();
    }

    /**
     * 获取斑马线样式值，它返回的是全局样式中斑马线填充样式的值，全局第{@code n}个填充返回 {@code n<<INDEX_FILL}，
     * 更多参考{@link Styles#addFill}
     *
     * @return 斑马线样式值
     */
    public int getZebraFillStyle() {
        if (zebraFillStyle < 0 && zebraFill != null) {
            this.zebraFillStyle = workbook.getStyles().addFill(zebraFill);
        }
        return zebraFillStyle;
    }

    /**
     * 设置默认斑马线，默认填充色HEX值为{@code E9EAEC}
     *
     * @return 当前工作表
     */
    public Sheet defaultZebraLine() {
        return setZebraLine(new Fill(PatternType.solid, new Color(233, 234, 236)));
    }

    /**
     * 获取当前工作表的表名
     *
     * <p>注意：仅返回实例化Worksheet时指定的表名或通过 {@link #setName}方法设置的表名，
     * 对于未指定的表名的工作表受分页和{@link Workbook#insertSheet(int, Sheet)}插入指定位置
     * 影响只能在最终执行输出时确定位置，在此之前表名均返回{@code null}</p>
     *
     * @return 外部指定的表名，未指定表名时返回{@code null}
     */
    public String getName() {
        return name;
    }

    /**
     * 设置工作表表名，使用Office打开文件时它将显示在底部的Tab栏
     *
     * <p>注意：内部不会检查重名，所以请在外部保证在一个工作薄下所有工作表名唯一，否则打开文件异常</p>
     *
     * @param name 工作表表名，最多31个字符超过时截取前31个字符
     * @return 当前工作表
     */
    public Sheet setName(String name) {
        if (name != null && name.length() > 31) {
            LOGGER.warn("The worksheet name is too long, maximum length of 31 characters. Currently {}", name.length());
            name = name.substring(0, 31);
        }
        this.name = name;
        return this;
    }

    /**
     * 获取批注 {@link Comments}
     *
     * @return 如果添加了批注则返回 {@code Comments}对象否则返回{@code null}
     */
    public Comments getComments() {
        if (comments != null && comments.id == 0) {
            comments.id = this.id;
        }
        return comments;
    }

    /**
     * 创建批注对象，一般由各工作表输出协议创建，外部用户勿用
     *
     * @return {@code Comments}实体，与工作表一一对应
     */
    public Comments createComments() {
        if (comments == null) {
            comments = new Comments(id, workbook != null ? workbook.getCreator() : null);
            // FIXME Removed at excel version 2013
            if (id > 0) {
                addRel(new Relationship("../drawings/vmlDrawing" + id + Const.Suffix.VML, Const.Relationship.VMLDRAWING));

                addRel(new Relationship("../comments" + id + Const.Suffix.XML, Const.Relationship.COMMENTS));
            }
        }
        return comments;
    }

    /**
     * 是否显示“网格线”
     *
     * @return true: 显示 false: 不显示
     */
    public boolean isShowGridLines() {
        return showGridLines == null || showGridLines;
    }

    /**
     * 设置显示“网格线”
     *
     * @return 当前工作表
     */
    public Sheet showGridLines() {
        this.showGridLines = true;
        return this;
    }

    /**
     * 设置隐藏“网格线”
     *
     * @return 当前工作表
     */
    public Sheet hideGridLines() {
        this.showGridLines = false;
        return this;
    }

    /**
     * 获取表头行高，默认20.5
     *
     * @return 表头行高
     */
    public double getHeaderRowHeight() {
        return headerRowHeight;
    }

    /**
     * 设置表头行高，其优化级低于{@link Column#setHeaderHeight}设置的值
     *
     * <p>可接受负数和零，负数等价与未设置默认行高为{@code 13.5}，零效果等价于隐藏，但不能通过右建“取消隐藏”</p>
     *
     * @param headerRowHeight 指定表头行高，建议表头行高比数据行大
     * @return 当前工作表
     */
    public Sheet setHeaderRowHeight(double headerRowHeight) {
        this.headerRowHeight = headerRowHeight;
        return this;
    }

    /**
     * 获取数据行高
     *
     * @return 数据行高，返回{@code null}时使用默认行高
     */
    public Double getRowHeight() {
        return rowHeight;
    }

    /**
     * 设置数据行高，未指定或负数时默认行高为{@code 13.5}
     *
     * @param rowHeight 指定数据行高
     * @return 当前工作表
     */
    public Sheet setRowHeight(double rowHeight) {
        this.rowHeight = rowHeight;
        return this;
    }

    /**
     * 获取工作表的起始行号(从1开始)，这里是行号也就是打开Excel左侧看到的行号，
     * 此行号将决定从哪一行开始写数据
     *
     * @return 起始行号
     */
    public int getStartRowIndex() {
        return startCoordinate != 0 ? (int) (Math.abs(startCoordinate) >>> 16) : 1;
    }

    /**
     * 获取工作表的起始列号(从1开始)，这里是行号也就是打开Excel顶部看到的列号(A)，
     * 此列号将决定从哪一列开始写数据
     *
     * @return 起始列号
     */
    public int getStartColIndex() {
        return startCoordinate != 0 ? (int) (Math.abs(startCoordinate) & 0x7FFF) : 1;
    }

    /**
     * 是否滚动到可视区，当起始行列不在{@code A1}时，如果返回{@code true}则打开Excel文件时
     * 自动将首行首列滚动左上角第一个位置，如果返回{@code false}时打开Excel文件左上角可视区为{@code A1}
     *
     * @return {@code true}将首行首列滚动到左上角第一个位置，否则{@code A1}将为左上角第一个位置
     */
    public boolean isScrollToVisibleArea() {
        return startCoordinate > 0;
    }

    /**
     * 指定起始行并将该行自动滚到窗口左上角，行号必须大于0
     *
     * @param startRowIndex 起始行号（从1开始）
     * @return 当前工作表
     * @deprecated 使用 {@link #setStartCoordinate(int)}替代
     */
    @Deprecated
    public Sheet setStartRowIndex(int startRowIndex) {
        return setStartRowIndex(startRowIndex, false);
    }

    /**
     * 指定起始行并设置是否将该行滚动到窗口左上角，行号必须大于0
     *
     * <p>默认情况下左上角一定是{@code A1}，如果{@code scrollToVisibleArea=true}则打开文件时{@code StartRowIndex}
     * 将会显示在窗口的第一行</p>
     *
     * @param startRowIndex       起始行号（从1开始）
     * @param scrollToVisibleArea 是否滚动起始行到窗口左上角
     * @return 当前工作表
     * @deprecated 使用 {@link #setStartCoordinate(int, boolean)}替代
     */
    @Deprecated
    public Sheet setStartRowIndex(int startRowIndex, boolean scrollToVisibleArea) {
        return setStartCoordinate(startRowIndex, 1, scrollToVisibleArea);
    }

    /**
     * 指定起始坐标
     *
     * @param coordinate 单元格位置字符串 {@code A1}
     * @return 当前工作表
     */
    public Sheet setStartCoordinate(String coordinate) {
        return setStartCoordinate(coordinate, false);
    }

    /**
     * 指定起始坐标
     *
     * @param coordinate 单元格位置字符串 {@code A1}
     * @return 当前工作表
     */
    public Sheet setStartCoordinate(String coordinate, boolean scrollToVisibleArea) {
        long f = coordinateToLong(coordinate);
        return setStartCoordinate((int) (f >> 16), (int) f & 0x7FFF, scrollToVisibleArea);
    }

    /**
     * 指定起始行，行号必须大于0
     *
     * @param startRowIndex 起始行号（从1开始）
     * @return 当前工作表
     */
    public Sheet setStartCoordinate(int startRowIndex) {
        return setStartCoordinate(startRowIndex, 1, false);
    }

    /**
     * 指定起始行，行号必须大于0
     *
     * @param startRowIndex 起始行号（从1开始）
     * @param scrollToVisibleArea 是否滚动起始行到窗口左上角
     * @return 当前工作表
     */
    public Sheet setStartCoordinate(int startRowIndex, boolean scrollToVisibleArea) {
        return setStartCoordinate(startRowIndex, 1, scrollToVisibleArea);
    }

    /**
     * 指定起始行号和列号，行号必须大于0
     *
     * @param startRowIndex 起始行号（从1开始）
     * @param startColIndex 起始列号（从1开始）
     * @return 当前工作表
     */
    public Sheet setStartCoordinate(int startRowIndex, int startColIndex) {
        return setStartCoordinate(startRowIndex, startColIndex, false);
    }

    /**
     * 指定起始行号和列号，行号必须大于0
     *
     * <p>默认情况下左上角一定是{@code A1}，如果{@code scrollToVisibleArea=true}则打开文件时{@code StartRowIndex}
     * 将会显示在窗口的第一行</p>
     *
     * @param startRowIndex       起始行号（从1开始）
     * @param startColIndex       起始列号（从1开始）
     * @param scrollToVisibleArea 是否滚动起始行到窗口左上角
     * @return 当前工作表
     */
    public Sheet setStartCoordinate(int startRowIndex, int startColIndex, boolean scrollToVisibleArea) {
        if (startRowIndex <= 0)
            throw new IndexOutOfBoundsException("The start row index must be greater than 0, current = " + startRowIndex);
        if (sheetWriter != null && sheetWriter.getRowLimit() <= startRowIndex)
            throw new IndexOutOfBoundsException("The start row index must be less than row-limit, current(" + startRowIndex + ") >= limit(" + sheetWriter.getRowLimit() + ")");
        if (startColIndex <= 0)
            throw new IndexOutOfBoundsException("The start col index must be greater than 0, current = " + startColIndex);
        if (sheetWriter != null && sheetWriter.getColumnLimit() <= startColIndex)
            throw new IndexOutOfBoundsException("The start col index must be less than col-limit, current(" + startColIndex + ") >= limit(" + sheetWriter.getColumnLimit() + ")");

        long coordinate = ((long) startRowIndex) << 16 | (startColIndex & 0x7FFF);
        this.startCoordinate = scrollToVisibleArea ? coordinate : -coordinate;
        return this;
    }

    /**
     * 获取表头，对于非外部传入的表头，只有要执行导出的时候通过行数据进行反射或读取Meta元数据获取，
     * 在此之前该接口将返回{@code null}
     *
     * @return 表头信息
     */
    public Column[] getColumns() {
        return columns;
    }

    /**
     * 添加进度观察者，在数据较大的导出过程中添加观察者打印进度可避免被误解为程序假死
     *
     * <pre>
     * new ListSheet&lt;&gt;().onProgress((sheet, row) -&gt; {
     *     System.out.println(sheet + " write " + row + " rows");
     * })</pre>
     *
     * @param progressConsumer 进度消费窗口
     * @return 当前工作表
     */
    public Sheet onProgress(BiConsumer<Sheet, Integer> progressConsumer) {
        this.progressConsumer = progressConsumer;
        return this;
    }

    /**
     * 获取进度观察者
     *
     * @return 如果设置了观察者则返回观察者否则返回 {@code null}
     */
    public BiConsumer<Sheet, Integer> getProgressConsumer() {
        return progressConsumer;
    }

    /**
     * 获取表头，子类覆写此方法创建表头
     *
     * @return 表头信息
     */
    protected Column[] getHeaderColumns() {
        if (!headerReady) {
            if (columns == null) {
                columns = new Column[0];
            }
        }
        return columns;
    }

    /**
     * 获取表头，Worksheet工作表输出协议调用此方法来获取表头信息
     *
     * <p>此方法先调用内部{@link #getHeaderColumns}获取基础信息，然后对其进行排序，列反转，
     * 合并等深加工处理</p>
     *
     * @return 加工好的表头
     */
    public Column[] getAndSortHeaderColumns() {
        if (!headerReady) {
            // 获取表头基础信息
            this.columns = getHeaderColumns();

            // Ready Flag
            headerReady |= (this.columns.length > 0);

            if (headerReady) {
                // 排序
                sortColumns(columns);

                // 计算每列在Excel中的列下标
                calculateRealColIndex();

                // 列反转，由于尾部Column包含必要的信息，多行表头时为方便获取主要信息这里进行一次反转
                reverseHeadColumn();

                // 合并，将相同列名的列进行合并
                mergeHeaderCellsIfEquals();

                // 重置通用属性
                resetCommonProperties(columns);

                // Check the limit of columns
                checkColumnLimit();
            }

            // Reset Row limit
//            this.rowLimit = sheetWriter.getRowLimit() - (nonHeader == 1 || columns.length == 0 ? 0 : columns[0].subColumnSize()) - getStartRowIndex() + 1

            // Mark ext-properties
            markExtProp();
        }
        return columns;
    }

    protected void resetCommonProperties(Column[] columns) {
        for (Column column : columns) {
            if (column == null) continue;
            if (column.styles == null) column.styles = workbook.getStyles();
            if (column.next != null) {
                for (Column col = column.next; col != null; col = col.next)
                    col.styles = workbook.getStyles();
            }

            // Column width
            if (column.getAutoSize() == 0 && autoSize > 0) {
                column.option |= autoSize << 1;
            }
        }
    }

    /**
     * 列排序，首先会根据用户指定的{@code colIndex}进行一次排序，未指定{@code colIndex}的列排在最后，
     * 然后将尾部没有{@code colIndex}的列插入到数组前方不连续的空白位，如果有重复的{@code colIndex}则按
     * 列在当前数组中的顺序依次排序
     *
     * <p>示例：现有A:1,B,C:4,D,E五列，其中A的{@code colIndex=1}，C的{@code colIndex=4}</p>
     *
     * <p>第一轮按{@code colIndex}排序后结果为 =&gt; {@code A:1,C:4,B,D,E}</p>
     *
     * <p>第二轮将尾部没有{@code colIndex}的BDE列插入到前方空白位，A在第1列它前方可以插入B，
     * A:1和C:4之间有2,3两个空白位，将DE分别插入到2，3位，现在结果为 =&gt; {@code B:0,A:1,D:2,E:3,C:4}</p>
     *
     * @param columns 表头信息
     */
    protected void sortColumns(Column[] columns) {
        if (columns.length <= 1) return;
        int j = 0;
        for (int i = 0; i < columns.length; i++) {
            if (columns[i].getTail().colIndex >= 0) {
                int n = search(columns, j, columns[i].getTail().colIndex);
                if (n < i) insert(columns, n, i);
                j++;
            }
        }
        // Finished
        if (j == columns.length) return;
        int n = columns[0].getTail().colIndex;
        for (int i = 0; i < columns.length && j < columns.length; ) {
            if (n > i) {
                for (int k = Math.min(n - i, columns.length - j); k > 0; k--, j++)
                    insert(columns, i++, j);
            } else i++;
            if (i < columns.length) n = columns[i].getTail().colIndex;
        }
    }

    protected int search(Column[] columns, int n, int k) {
        int i = 0;
        for (; i < n && columns[i].getTail().colIndex <= k; i++) ;
        return i;
    }

    protected void insert(Column[] columns, int n, int k) {
        Column t = columns[k];
        System.arraycopy(columns, n, columns, n + 1, k - n);
        columns[n] = t;
    }

    /**
     * 计算列的实际下标，Excel下标从1开始，计算后的值将重置{@link Column#realColIndex}属性，
     * 该属性将最终输出到Excel文件{@code col}属性中
     */
    protected void calculateRealColIndex() {
        int startColIndex = getStartColIndex();
        for (int i = 0; i < columns.length; i++) {
            Column hc = columns[i].getTail();
            hc.realColIndex = hc.colIndex;
            if (i > 0 && columns[i - 1].realColIndex >= hc.realColIndex)
                hc.realColIndex = columns[i - 1].realColIndex + 1;
            else if (hc.realColIndex <= i) hc.realColIndex = i + startColIndex;
            else hc.realColIndex = hc.colIndex + startColIndex;

            if (hc.prev != null) {
                for (Column col = hc.prev; col != null; col = col.prev)
                    col.realColIndex = hc.realColIndex;
            }
        }
    }

    /**
     * 设置表头，无数据时依然会导出该表头
     *
     * @param columns 表头数组
     * @return 当前工作表
     */
    public Sheet setColumns(final Column ... columns) {
        this.columns = columns;
        return this;
    }

    /**
     * 设置表头，无数据时依然会导出该表头
     *
     * @param columns 表头数组
     * @return 当前工作表
     */
    public Sheet setColumns(List<Column> columns) {
        if (columns != null && !columns.isEmpty()) {
            this.columns = new Column[columns.size()];
            columns.toArray(this.columns);
        }
        return this;
    }

    /**
     * 获取水印
     *
     * @return 水印对象 {@link WaterMark}
     */
    public WaterMark getWaterMark() {
        return waterMark;
    }

    /**
     * 设置水印，优先级高于Workbook中的全局水印
     *
     * @param waterMark 水印对象 {@link WaterMark}
     * @return 当前工作表
     */
    public Sheet setWaterMark(WaterMark waterMark) {
        this.waterMark = waterMark;
        return this;
    }

    /**
     * 工作表是否隐藏
     *
     * @return true: 隐藏, false: 显示
     */
    public boolean isHidden() {
        return hidden;
    }

    /**
     * 隐藏工作表
     *
     * @return 当前工作表
     */
    public Sheet hidden() {
        this.hidden = true;
        return this;
    }

    /**
     * 获取强制导出标识，只对{@link ListSheet}生效，用于
     *
     * @return 1: 强制导出 其它值均表示不强制导出
     */
    public int getForceExport() {
        return 0;
    }

    /**
     * 回闭连接，回收资源，删除临时文件等
     *
     * @throws IOException if I/O error occur
     */
    public void close() throws IOException {
        if (sheetWriter != null) {
            sheetWriter.close();
        }
    }

    /**
     * 落盘，将工作表写到指定路径
     *
     * @param path 指定保存路径
     * @throws IOException if I/O error occur
     */
    @Override
    public void writeTo(Path path) throws IOException {
        if (sheetWriter == null) {
            throw new ExcelWriteException("Worksheet writer is not instanced.");
        }
        if (!headerReady) {
            getAndSortHeaderColumns();
        }
        if (rowBlock == null) {
            rowBlock = new RowBlock(getRowBlockSize());
        }
        // 自动分页的Sheet可复用RowBlock
        else rowBlock.reopen();

        if (!copySheet) {
            paging();
        }

        sheetWriter.writeTo(path);
    }

    /**
     * 分批拉取数据
     */
    protected void paging() { }

    /**
     * 添加关联，当工作表需要引入其它资源时必须将其添加进关联关系中，关联关系由{@link RelManager}管理。
     *
     * <p>例如：向工作表添加图片时，图片由media统一存放，工作表中只需要加入图片的关联关系，通过rId值找到图片。
     * 除了图片外，像批注，公式，图表等都属于外部资源</p>
     *
     * @param rel 关联关系 {@link Relationship}
     * @return 当前工作表
     */
    public Sheet addRel(Relationship rel) {
        relManager.add(rel);
        return this;
    }

    /**
     * 通过相对位置模糊匹配查找关联关系
     *
     * @param key 要查询的关联key
     * @return 第一个匹配的关联，如果未匹配则返回{@code null}
     */
    public Relationship findRel(String key) {
        return relManager.likeByTarget(key);
    }

    /**
     * 获取当前工作表的关系管理器
     *
     * @return {@link RelManager}
     */
    public RelManager getRelManager() {
        return relManager;
    }

    /**
     * 获取当前工作表的文件名
     *
     * @return 工作表文件名
     */
    public String getFileName() {
        return "sheet" + id + sheetWriter.getFileSuffix();
    }

    /**
     * 设置统一的表头样式
     *
     * @param font   字体
     * @param fill   填充色
     * @param border 边框
     * @return 当前工作表
     * @deprecated 可能因为Style未初始化出现 {@code NPE}，目前最可靠的只有{@link #setHeadStyle(int)}
     */
    @Deprecated
    public Sheet setHeadStyle(Font font, Fill fill, Border border) {
        return setHeadStyle(null, font, fill, border, Verticals.CENTER, Horizontals.CENTER);
    }

    /**
     * 设置统一的表头样式
     *
     * @param font       字体
     * @param fill       填充色
     * @param border     边框
     * @param vertical   垂直对齐
     * @param horizontal 水平对齐
     * @return 当前工作表
     * @deprecated 可能因为Style未初始化出现 {@code NPE}，目前最可靠的只有{@link #setHeadStyle(int)}
     */
    @Deprecated
    public Sheet setHeadStyle(Font font, Fill fill, Border border, int vertical, int horizontal) {
        return setHeadStyle(null, font, fill, border, vertical, horizontal);
    }

    /**
     * 设置统一的表头样式
     *
     * @param numFmt     格式化
     * @param font       字体
     * @param fill       填充色
     * @param border     边框
     * @param vertical   垂直对齐
     * @param horizontal 水平对齐
     * @return 当前工作表
     * @deprecated 可能因为Style未初始化出现 {@code NPE}，目前最可靠的只有{@link #setHeadStyle(int)}
     */
    @Deprecated
    public Sheet setHeadStyle(NumFmt numFmt, Font font, Fill fill, Border border, int vertical, int horizontal) {
        Styles styles = workbook.getStyles();
        headStyle = (numFmt != null ? styles.addNumFmt(numFmt) : 0)
            | (font != null ? styles.addFont(font) : 0)
            | (fill != null ? styles.addFill(fill) : 0)
            | (border != null ? styles.addBorder(border) : 0)
            | vertical
            | horizontal;
        headStyleIndex = styles.of(headStyle);
        return this;
    }

    /**
     * 设置统一的表头样式值
     *
     * @param style 样式值，0表示默认样式
     * @return 当前工作表
     */
    public Sheet setHeadStyle(int style) {
        headStyle = style;
        headStyleIndex = workbook.getStyles().of(style);
        return this;
    }

    /**
     * 设置统一的表头样式索引
     *
     * @param styleIndex 样式索引，索引从0开始，负数表示未设置样式
     * @return 当前工作表
     */
    public Sheet setHeadStyleIndex(int styleIndex) {
        headStyleIndex = styleIndex;
        headStyle = workbook.getStyles().getStyleByIndex(styleIndex);
        return this;
    }

    /**
     * 获取统一的表头样式值
     *
     * @return 样式值，0表示默认样式
     */
    public int getHeadStyle() {
        return headStyle;
    }

    /**
     * 获取统一的表头样式索引
     *
     * @return 样式索引，索引从0开始，负数表示未设置样式
     */
    public int getHeadStyleIndex() {
        return headStyleIndex;
    }

    /**
     * 使用默认样式并修改文字颜色和充填色创建统一表头样式
     *
     * @param fontColor   文字颜色，可以使用{@link java.awt.Color}中定义的颜色名或者Hex值
     * @param fillBgColor 充填色，可以使用{@link java.awt.Color}中定义的颜色名或者Hex值
     * @return 样式值
     */
    public int buildHeadStyle(String fontColor, String fillBgColor) {
        Styles styles = workbook.getStyles();
        return styles.addFont(new Font("宋体", 12, Font.Style.BOLD, Styles.toColor(fontColor)))
                | styles.addFill(Fill.parse(fillBgColor))
                | styles.addBorder(new Border(BorderStyle.THIN, new Color(191, 191, 191)))
                | Verticals.CENTER
                | Horizontals.CENTER;
    }

    /**
     * 获取默认的表头样式值
     *
     * @return 样式值
     */
    public int defaultHeadStyle() {
        return headStyle != 0 ? headStyle : (headStyle = this.buildHeadStyle("black", "#E9EAEC"));
    }

    /**
     * 获取默认的表头样式索引
     *
     * @return 样式索引
     */
    public int defaultHeadStyleIndex() {
        if (headStyleIndex == -1) {
            setHeadStyle(this.buildHeadStyle("black", "#E9EAEC"));
        }
        return headStyleIndex;
    }

    /**
     * 获取已写入的数据行数，这里不包含表头行
     *
     * <p>注意：由于数据行经由{@link RowBlock}行块统一处理，所以这里的已写入只表示写入到
     * {@link RowBlock}行块的数据并非实际已导出的行数，精准的已导出行可以通过{@link Workbook#onProgress}监听获取</p>
     *
     * @return 写入的数据行数，-1表示不确定
     */
    public int size() {
        return !shouldClose ? rows : -1;
    }

    /**
     * 获取下一段{@link RowBlock}行块数据，工作表输出协议通过此方法循环获取行数据并落盘，
     * 行块被设计为一个滑行窗口，下游输出协议只能获取一个窗口的数据默认包含32行。
     *
     * @return 行块
     */
    public RowBlock nextBlock() {
        // clear first
        rowBlock.clear();

        if (columns.length > 0) {
            resetBlockData();
        }

        return rowBlock.flip();
    }

    /**
     * 获取{@link RowBlock}行块的大小，创建行块时会调用此方法获取行块大小，
     * 子类可覆写该方法指定其它值
     *
     * @return 行块大小
     */
    public int getRowBlockSize() {
        return ROW_BLOCK_SIZE;
    }

    /**
     * 当输出协议写完sheetData时调用
     *
     * @param total 已写数据行
     */
    public void afterSheetDataWriter(int total) { }

    /**
     * 当输出协议输出完成时调用此方法输出关联
     *
     * @param workSheetPath 当前工作表保存路径
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
     * 当数据行超过工作表限制时触发分页，复制得到新的工作表命名为原工作表名+页码数
     *
     * @return 复制工作表的表名
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

    /**
     * 深拷贝当前工作表，分页时使用
     *
     * @return 当前工作表的副本
     */
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
                } catch (NoSuchMethodException | IllegalAccessException | InstantiationException | InvocationTargetException e2) { }
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
        int a = columns.length > 0 ? columns[columns.length - 1].getRealColIndex() : 0
            , b = sheetWriter.getColumnLimit();
        if (a > b) {
            throw new TooManyColumnsException(a, b);
        } else if (nonHeader == -1 && headerReady) {
            boolean noneHeader = columns == null || columns.length == 0;
            if (!noneHeader) {
                int n = 0;
                for (Column column : columns) {
                    if (isEmpty(column.name) && isEmpty(column.key)) n++;
                }
                noneHeader = n == columns.length;
            }
            if (noneHeader) {
                if (rows > 0) rows--;
                ignoreHeader();
            } else this.nonHeader = 0;
        }
    }

    /**
     * 是否包含表头
     *
     * @return true: 已收集表头
     */
    public boolean hasHeaderColumns() {
        return columns != null && columns.length > 0;
    }

    /**
     * 列下标转为Excel列标识，Excel列标识由大写字母{@code A-Z}组合，{@code Z}后为{@code AA}如此循环，最大下标{@code XFD}
     *
     * <blockquote><pre>
     * 数字    | Excel列
     * -------+---------
     * 1      | A
     * 10     | J
     * 26     | Z
     * 27     | AA
     * 28     | AB
     * 53     | BA
     * 16_384 | XFD
     * </pre></blockquote>
     *
     * @param n 列下标
     * @return Excel列标识
     */
    public static char[] int2Col(int n) {
        char[] c;
        char A = 'A';
        if (n <= 26) {
            c = tmpBuf[0];
            c[0] = (char) (n - 1 + A);
        } else if (n <= 702) {
            int t = n / 26, w = n % 26;
            if (w == 0) {
                t--;
                w = 26;
            }
            c = tmpBuf[1];
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
            c = tmpBuf[2];
            c[0] = (char) (t - 1 + A);
            c[1] = (char) (m - 1 + A);
            c[2] = (char) (w - 1 + A);
        }
        return c;
    }

    private static final char[][] tmpBuf = new char[][]{ {65}, {65, 65}, {65, 65, 65} };

    /**
     * 将行列坐标转换为 Excel 样式的单元格地址
     *
     * @param row 行号，从{@code 1}开始
     * @param col 列号，从{@code 1}开始
     * @return Excel 样式的单元格地址，例如{@code A1}、{@code B2}等
     */
    public static String toCoordinate(int row, int col) {
        char[] cols = int2Col(col);
        char[] chars = new char[cols.length + stringSize(row)];
        System.arraycopy(cols, 0, chars, 0, cols.length);
        getChars(row, chars.length, chars);
        return new String(chars);
    }

    /**
     * 忽略表头，调用此方法后表头将不会输出到Excel中，注意这里不是隐藏
     *
     * @return 当前工作表
     */
    public Sheet ignoreHeader() {
        this.nonHeader = 1;
        return this;
    }

    /**
     * 获取忽略表头标识
     *
     * @return -1 未设置, 0 输出表头, 1 忽略表头
     */
    public int getNonHeader() {
        return nonHeader;
    }

    /**
     * 获取工作表数据行上限，超过上限时触发分页，默认情况下此值由各输出协议决定
     *
     * @return 数据行上限
     */
    protected int getRowLimit() {
        return rowLimit > 0 ? rowLimit : (rowLimit = sheetWriter.getRowLimit() - (nonHeader == 1 || columns.length == 0 ? 0 : columns[0].subColumnSize()) - getStartRowIndex() + 1);
    }

    /**
     * 添加扩展参数，key存在时覆盖原值
     *
     * @param key   扩展参数Key
     * @param value 值
     * @return 当前工作表
     */
    public Sheet putExtProp(String key, Object value) {
        extProp.put(key, value);
        return this;
    }

    /**
     * 添加扩展参数，如果key不存在则添加，存在时忽略
     *
     * @param key   扩展参数Key
     * @param value 值
     * @return 当前工作表
     */
    public Sheet putExtPropIfAbsent(String key, Object value) {
        extProp.putIfAbsent(key, value);
        return this;
    }

    /**
     * 批量添加扩展参数，key存在时覆盖原值
     *
     * @param m 扩展参数
     * @return 当前工作表
     */
    public Sheet putAllExtProp(Map<String, Object> m) {
        if (m != null) extProp.putAll(m);
        return this;
    }

    /**
     * 根据key获取扩展参数
     *
     * @param key 扩展参数Key
     * @return key对应的扩展参数，不存在时返回{@code null}
     */
    public Object getExtPropValue(String key) {
        return extProp.get(key);
    }

    /**
     * 获取所有扩展参数
     *
     * <p>注意：对返回值进行修改可能会影响原始值，添加/删除无影响</p>
     *
     * @return 扩展参数
     */
    public Map<String, Object> getExtPropAsMap() {
        return new HashMap<>(extProp);
    }

    /**
     * 标记扩展参数（非必要操作）
     */
    protected void markExtProp() {
        // Mark Freeze Panes
        extPropMark |= getExtPropValue(Const.ExtendPropertyKey.FREEZE) != null ? 1 : 0;
        // Mark global style design
        extPropMark |= getExtPropValue(Const.ExtendPropertyKey.STYLE_DESIGN) != null ? 1 << 1 : 0;
        // Mark global merged cells
        extPropMark |= getExtPropValue(Const.ExtendPropertyKey.MERGE_CELLS) != null ? 1 << 2 : 0;
        // Mark auto-filter
        extPropMark |= getExtPropValue(Const.ExtendPropertyKey.AUTO_FILTER) != null ? 1 << 3 : 0;
        // Mark data-validation
        extPropMark |= getExtPropValue(Const.ExtendPropertyKey.DATA_VALIDATION) != null ? 1 << 4 : 0;
        // Mark Zoom-Scale
        extPropMark |= getExtPropValue(Const.ExtendPropertyKey.ZOOM_SCALE) != null ? 1 << 5 : 0;
    }

    /**
     * 反转表头
     */
    protected void reverseHeadColumn() {
        if (!headerReady) this.columns = getHeaderColumns();
        if (columns == null || columns.length == 0) return;

        // Count the number of sub-columns
        int[] lenArray = new int[columns.length];
        int maxSubColumnSize = 1;
        for (int i = 0, a, len = columns.length; i < len; i++) {
            Column col = columns[i];
            a = col.subColumnSize();
            lenArray[i] = a;
            if (a > maxSubColumnSize) {
                maxSubColumnSize = a;
            }
        }
        // Single header column
        if (maxSubColumnSize == 1) return;

        // Reverse and fill empty column
        for (int i = 0, len = columns.length; i < len; i++) {
            Column col = columns[i];
            // Reverse header to tail
            if (col.tail != null) {
                Column head = col.tail, tmp = head.prev;
                head.tail = null; head.prev = null; head.next = null;
                // Switch prev and next point
                for (; tmp != null; ) {
                    Column ptmp = tmp.prev;
                    tmp.tail = null; tmp.prev = null; tmp.next = null;
                    head.addSubColumn(tmp);
                    tmp = ptmp;
                }
                head.prev = null;
                if (head.tail != null) head.tail.next = null;
                columns[i] = head;
                col = head;
            }
            // Fill empty column
            if (lenArray[i] < maxSubColumnSize) {
                for (int k = lenArray[i]; k < maxSubColumnSize; k++) {
                    Column sub = new Column().setColIndex(col.colIndex);
                    sub.realColIndex = col.realColIndex;
                    col.addSubColumn(sub);
                }
            }
        }
    }

    /**
     * 合并表头
     */
    protected void mergeHeaderCellsIfEquals() {
        int x = columns.length, y = x > 0 ? columns[0].subColumnSize() : 0, n = x * y;
        // Single header column
        if (y <= 1) return;

        Column[] array = new Column[n];
        for (int i = 0; i < x; i++) {
            System.arraycopy(columns[i].toArray(), 0, array, y * i, y);
        }

        // Mark as 1 if visited
        int[] marks = new int[n];

        int fc = 0, fr = 0, lc = 0, lr = 0;
        List<Dimension> mergeCells = new ArrayList<>(), _tmpCells = new ArrayList<>();
        for (int i = 0; i < n; i++) {
            // Skip if marked
            if (marks[i] == 1) continue;

            Column col = array[i];
            marks[i] = 1;
            if (isEmpty(col.name)) {
                continue;
            }
            int a = 0;
            if (i + y < n && col.name.equals(array[i + y].name)) {
                fc = i / y; fr = i % y; lc = fc + 1; lr = fr;
                a = 1;
                marks[i + y] = 1;
                for (int c; (c = i + (y * (a + 1))) < n; a++) {
                    if (col.name.equals(array[c].name)) {
                        lc++;
                        marks[c] = 1;
                    } else break;
                }
            }
            int tail = i / y * y + y, r;
            if (i + 1 < tail && (col.name.equals(array[i + 1].name) || isEmpty(array[i + 1].name))) {
                r = i + 1;
                marks[r] = 1;
                fc = i / y; fr = i % y; lc = fc + a; lr = fr;
                A: for (; r < tail; r++) {
                    for (int k = 0; k <= a; k++) {
                        if (!col.name.equals(array[r + k * y].name) && isNotEmpty(array[r + k * y].name))
                            break A;
                    }
                    for (int k = 0; k <= a; k++) {
                        marks[r + k * y] = 1;
                    }
                    lr++;
                }
                i = r - 1;
            }
            // Add merged cells
            if (fc < lc || fr < lr) {
                mergeCells.add(new Dimension(y - lr, (short) array[y - lr - 1 + fc * y].realColIndex, y - fr, (short) array[y - fr - 1 + lc * y].realColIndex));
                _tmpCells.add(new Dimension(y - lr, (short) (fc + 1), y - fr, (short) (lc + 1)));
                // Reset
                fc = lc; fr = lr;
            }
        }

        // Put merged-cells into ext-properties
        if (!mergeCells.isEmpty()) {
            for (Dimension dim : _tmpCells) {
                Column col = array[(dim.firstColumn - 1) * y + (y - dim.lastRow)];
                Comment headerComment = col.headerComment;
                Column tmp = new Column().from(col);

                // Clear name in merged cols range
                for (int m = dim.firstColumn - 1; m < dim.lastColumn; m++) {
                    for (int o = y - dim.firstRow; o >= y - dim.lastRow; o--) {
                        Column currentCol = array[m * y + o];
                        currentCol.name = null;
                        currentCol.key = null;
                        if (currentCol.headerComment != null) {
                            if (headerComment == null) {
                                headerComment = currentCol.headerComment;
                            }
                            currentCol.headerComment = null;
                        }
                    }
                }

                // Copy last col's name into first col
                Column lastCol = array[(dim.firstColumn - 1) * y + (y - dim.firstRow)];
                lastCol.from(tmp);
                lastCol.headerComment = headerComment;
            }

            if (getStartRowIndex() > 1) {
                List<Dimension> tmp = new ArrayList<>();
                for (Dimension dim : mergeCells) {
                    tmp.add(new Dimension(dim.firstRow + getStartRowIndex() - 1, dim.firstColumn, dim.lastRow + getStartRowIndex() - 1, dim.lastColumn));
                }
                mergeCells = tmp;
            }

            @SuppressWarnings("unchecked")
            List<Dimension> existsMergeCells = (List<Dimension>) getExtPropValue(Const.ExtendPropertyKey.MERGE_CELLS);
            if (existsMergeCells != null && !existsMergeCells.isEmpty()) existsMergeCells.addAll(mergeCells);
            else putExtProp(Const.ExtendPropertyKey.MERGE_CELLS, mergeCells);
        }
    }

    ////////////////////////////Abstract function\\\\\\\\\\\\\\\\\\\\\\\\\\\

    /**
     * 你需要覆写本方法获取行数据并将Java对象或自定义行数据通过{@link ICellValueAndStyle#reset}转换器将数据转换为输出协议允许的结构，
     * {@code ICellValueAndStyle#reset}方法除转换数据外还添单元格样式
     */
    protected abstract void resetBlockData();
}

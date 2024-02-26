/*
 * Copyright (c) 2017-2023, guanquan.wang@yandex.com All Rights Reserved.
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
import org.ttzero.excel.entity.style.NumFmt;
import org.ttzero.excel.entity.style.Styles;
import org.ttzero.excel.manager.Const;
import org.ttzero.excel.reader.Cell;
import org.ttzero.excel.reader.Col;
import org.ttzero.excel.reader.Dimension;
import org.ttzero.excel.reader.Drawings;
import org.ttzero.excel.reader.ExcelReader;
import org.ttzero.excel.reader.FullSheet;
import org.ttzero.excel.util.DateUtil;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Path;
import java.util.Comparator;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

/**
 * 模板工作表，它支持指定一个已有的Excel文件作为模板导出，{@code TemplateSheet}将复制
 * 模板工作表的样式并替换占位符，同时{@code TemplateSheet}也可以和其它{@code Sheet}共用，
 * 意味着可以添加多个模板工作表和普通工作表。需要注意的是多个模板可能产生重复的工作表名称，所以需要外部指定不同的名称以免
 * 打开文件异常
 *
 * <p>创建模板工作表需要指定模板文件，它可以是本地文件也可是输入流{@code InputStream}，支持的类型包含{@code xls}
 * 和{@code xlsx}两种格式，除模板文件外还需要指定Excel中的某个{@code Worksheet}，
 * 未指定工作表时默认以第一个工作表做为模板，{@code TemplateSheet}工作表导出时不受{@code ExcelColumn}注解限制，
 * 导出的数据范围由模板内占位符决定</p>
 *
 * <p>默认占位符为一对关闭的大括号{@code ‘${key}’}，</p>
 *
 * <p>考虑到模板工作表的复杂性暂时不支持数据切片，数据必须在初始化时设置，换句话说模板工作表只适用于少量数据</p>
 *
 * <blockquote><pre>
 * new Workbook("模板测试")
 *     .addSheet(new TemplateSheet(Paths.get("./template.xlsx")).setData(data)) // &lt;- 模板工作表
 *     .addSheet(new ListSheet&lt;&gt;()) // &lt;- 普通对象数组工作表
 *     .writeTo("/tmp/");</pre></blockquote>
 *
 * @author guanquan.wang at 2023-12-01 15:10
 */
public class TemplateSheet extends Sheet {
    /**
     * 读取模板用
     */
    protected ExcelReader reader;
    /**
     * 源工作表索引
     */
    protected int originalSheetIndex;
    /**
     * 行数据迭代器
     */
    protected Iterator<org.ttzero.excel.reader.Row> rowIterator;
    /**
     * 样式映射，缓存源样式索引映射到目标样式索引
     */
    protected Map<Integer, Integer> styleMap;

    /**
     * 实例化模板工作表，默认以第一个工作表做为模板
     *
     * @param templatePath 模板路径
     * @throws IOException 文件不存在或读取模板异常
     */
    public TemplateSheet(Path templatePath) throws IOException {
        this(templatePath, 0);
    }

    /**
     * 实例化模板工作表，默认以第一个工作表做为模板
     *
     * @param name         指定工作表名称
     * @param templatePath 模板路径
     * @throws IOException 文件不存在或读取模板异常
     */
    public TemplateSheet(String name, Path templatePath) throws IOException {
        this(name, templatePath, 0);
    }

    /**
     * 实例化模板工作表并指定模板工作表索引，如果指定索引超过模板Excel中包含的工作表数量则抛异常
     *
     * @param templatePath       模板路径
     * @param originalSheetIndex 指定源工作表索引（从0开始）
     * @throws IOException 文件不存在或读取模板异常
     */
    public TemplateSheet(Path templatePath, int originalSheetIndex) throws IOException {
        this(null, templatePath, originalSheetIndex);
    }

    /**
     * 实例化模板工作表并指定模板工作表索引，如果指定索引超过模板Excel中包含的工作表数量则抛异常
     *
     * @param name               指定工作表名称
     * @param templatePath       模板路径
     * @param originalSheetIndex 指定源工作表索引（从0开始）
     * @throws IOException 文件不存在或读取模板异常
     */
    public TemplateSheet(String name, Path templatePath, int originalSheetIndex) throws IOException {
        this.name = name;
        this.reader = ExcelReader.read(templatePath);
        if (reader.getSheetCount() < originalSheetIndex)
            throw new IOException("Original sheet index [" + originalSheetIndex + "] does not exist in template file.");
    }

    /**
     * 实例化模板工作表并指定模板工作表名，如果指定源工作表不存在则抛异常
     *
     * @param templatePath      模板路径
     * @param originalSheetName 指定源工作表名
     * @throws IOException 文件不存在或读取模板异常
     */
    public TemplateSheet(Path templatePath, String originalSheetName) throws IOException {
        this(null, templatePath, originalSheetName);
    }

    /**
     * 实例化模板工作表并指定模板工作表名，如果指定源工作表不存在则抛异常
     *
     * @param name              指定工作表名称
     * @param templatePath      模板路径
     * @param originalSheetName 指定源工作表名
     * @throws IOException 文件不存在或读取模板异常
     */
    public TemplateSheet(String name, Path templatePath, String originalSheetName) throws IOException {
        this.name = name;
        this.reader = ExcelReader.read(templatePath);
        org.ttzero.excel.reader.Sheet[] sheets = reader.all();
        int index = 0;
        for (; index < sheets.length && !originalSheetName.equals(sheets[index].getName()); index++) ;
        if (index >= sheets.length)
            throw new IOException("The specified sheet [" + originalSheetName + "] does not exist in template file.");
        originalSheetIndex = index;
    }

    /**
     * 实例化模板工作表，默认以第一个工作表做为模板
     *
     * @param templateStream 模板输入流
     * @throws IOException 读取模板异常
     */
    public TemplateSheet(InputStream templateStream) throws IOException {
        this(templateStream, 0);
    }

    /**
     * 实例化模板工作表，默认以第一个工作表做为模板
     *
     * @param name           设置工作表名
     * @param templateStream 模板输入流
     * @throws IOException 读取模板异常
     */
    public TemplateSheet(String name, InputStream templateStream) throws IOException {
        this(name, templateStream, 0);
    }

    /**
     * 实例化模板工作表并指定模板工作表索引，如果指定索引超过模板Excel中包含的工作表数量则抛异常
     *
     * @param templateStream     模板输入流
     * @param originalSheetIndex 指定源工作表索引
     * @throws IOException 读取模板异常
     */
    public TemplateSheet(InputStream templateStream, int originalSheetIndex) throws IOException {
        this(null, templateStream, originalSheetIndex);
    }

    /**
     * 实例化模板工作表并指定模板工作表名，如果指定源工作表不存在则抛异常
     *
     * @param templateStream    模板输入流
     * @param originalSheetName 指定源工作表名
     * @throws IOException 读取模板异常
     */
    public TemplateSheet(InputStream templateStream, String originalSheetName) throws IOException {
        this(null, templateStream, originalSheetName);
    }

    /**
     * 实例化模板工作表并指定模板工作表索引，如果指定索引超过模板Excel中包含的工作表数量则抛异常
     *
     * @param name               设置工作表名
     * @param templateStream     模板输入流
     * @param originalSheetIndex 指定源工作表索引
     * @throws IOException 读取模板异常
     */
    public TemplateSheet(String name, InputStream templateStream, int originalSheetIndex) throws IOException {
        this.name = name;
        this.reader = ExcelReader.read(templateStream);
        if (reader.getSheetCount() < originalSheetIndex)
            throw new IOException("Original sheet index [" + originalSheetIndex + "] does not exist in template file.");
    }

    /**
     * 实例化模板工作表并指定模板工作表名，如果指定源工作表不存在则抛异常
     *
     * @param name              设置工作表名
     * @param templateStream    模板输入流
     * @param originalSheetName 指定源工作表名
     * @throws IOException 读取模板异常
     */
    public TemplateSheet(String name, InputStream templateStream, String originalSheetName) throws IOException {
        this.name = name;
        this.reader = ExcelReader.read(templateStream);
        org.ttzero.excel.reader.Sheet[] sheets = reader.all();
        int index = 0;
        for (; index < sheets.length && !originalSheetName.equals(sheets[index++].getName()); ) ;
        if (index >= sheets.length)
            throw new IOException("The specified sheet [" + originalSheetName + "] does not exist in template file.");
        originalSheetIndex = index;
    }

    /**
     * 获取下一段{@link RowBlock}行块数据，工作表输出协议通过此方法循环获取行数据并落盘，
     * 行块被设计为一个滑行窗口，下游输出协议只能获取一个窗口的数据默认包含32行。
     *
     * @return 行块
     */
    public RowBlock nextBlock() {
        // 清除数据（仅重置下标）
        rowBlock.clear();

        // 装载数据（这里不需要判断是否有表头，模板是没有表头的）
        resetBlockData();

        // 使其可读
        return rowBlock.flip();
    }

    @Override
    public Column[] getAndSortHeaderColumns() {
        if (!headerReady) {
            // 解析模板工作表并复制信息到当前工作表中
            int size = init();
            if (size <= 0) columns = new Column[0];
            else {
                // 排序
                sortColumns(columns);
                // 计算每列在Excel中的列下标
                calculateRealColIndex();
                // 重置通用属性
                resetCommonProperties(columns);
            }
            // Mark ext-properties
            markExtProp();

            headerReady = true;
        }
        return columns;
    }

    /**
     * 读取模板头信息并复杂到当前工作表
     *
     * @return 列的个数
     */
    protected int init() {
        // 加载模板工作表
        FullSheet sheet = reader.sheet(originalSheetIndex).asFullSheet();

        // 冻结,直接复制不需要计算移动
        Panes panes = sheet.getFreezePanes();
        if (panes != null) putExtProp(Const.ExtendPropertyKey.FREEZE, panes);

        // TODO 合并（较为复杂不能简单复制，需要计算中间插入或扣除的行）
        List<Dimension> mergeCells = sheet.getMergeCells();
        if (mergeCells != null) putExtProp(Const.ExtendPropertyKey.MERGE_CELLS, mergeCells);

        // 过滤
        Dimension autoFilter = sheet.getFilter();
        if (autoFilter != null) putExtProp(Const.ExtendPropertyKey.AUTO_FILTER, autoFilter);

        // 是否显示网格线
        this.showGridLines = sheet.isShowGridLines();

        // 获取列属性
        int len = 0;
        List<Col> cols = sheet.getCols();
        if (cols != null && !cols.isEmpty()) {
            cols.sort(Comparator.comparingInt(a -> a.max));
            // 创建列
            len = cols.get(cols.size() - 1).max;
            int i = 0;
            columns = new Column[len];
            for (Col col : cols) {
                if (i + 1 < col.min) {
                    for (int a = i + 1; a < col.min; a++) {
                        Column c = new Column();
                        c.colIndex = a - 1;
                        columns[i++] = c;
                    }
                }
                for (int a = col.min; a <= col.max; a++) {
                    Column c = new Column();
                    c.width = col.width;
                    c.colIndex = a - 1;
                    if (col.hidden) c.hide();
                    columns[i++] = c;
                }
            }
        }
        // 忽略表头输出
        super.ignoreHeader();

        // 预置列宽
        double defaultColWidth = sheet.getDefaultColWidth(), defaultRowHeight = sheet.getDefaultRowHeight();
        if (defaultColWidth >= 0) putExtProp("defaultColWidth", defaultColWidth);
        if (defaultRowHeight >= 0) putExtProp("defaultRowHeight", defaultRowHeight);

        // 图片
        List<Drawings.Picture> pictures = sheet.listPictures();
        // FIXME 其它图片支持
        if (pictures != null && !pictures.isEmpty()) {
            for (Drawings.Picture p : pictures) {
                if (p.isBackground()) setWaterMark(WaterMark.of(p.getLocalPath()));
            }
        }

        // 初始化行迭代器
        rowIterator = sheet.iterator();

        styleMap = new HashMap<>();

        return len;
    }

    @Override
    protected void resetBlockData() {
        int len, n = 0, limit = getRowLimit();
        // 模板文件样式
        Styles styles0 = reader.getStyles(), styles = workbook.getStyles();

        for (int rbs = rowBlock.capacity(); n++ < rbs && rows < limit && rowIterator.hasNext(); rows++) {
            Row row = rowBlock.next();
            org.ttzero.excel.reader.Row row0 = rowIterator.next();
            // 设置行号
            row.index = rows = row0.getRowNum() - 1;
            // 设置行高
            row.height = row0.getHeight();
            // 设置行是否隐藏
            row.hidden = row0.isHidden();
            // 空行特殊处理（lc-fc=-1)
            len = Math.max(row0.getLastColumnIndex() - row0.getFirstColumnIndex(), 0);
            Cell[] cells = row.realloc(len);
            for (int i = 0; i < len; i++) {
                // clear cells
                Cell cell = cells[i], cell0 = row0.getCell(i);
                cell.clear();

                // 复制数据
                switch (row0.getCellType(cell0)) {
                    case STRING:  cell.setString(row0.getString(cell0));                                break;
                    case LONG:    cell.setLong(row0.getLong(cell0));                                    break;
                    case INTEGER: cell.setInt(row0.getInt(cell0));                                      break;
                    case DECIMAL: cell.setDecimal(row0.getDecimal(cell0));                              break;
                    case DOUBLE:  cell.setDouble(row0.getDouble(cell0));                                break;
                    case DATE:    cell.setDateTime(DateUtil.toDateTimeValue(row0.getTimestamp(cell0))); break;
                    case BOOLEAN: cell.setBool(row0.getBoolean(cell0));                                 break;
                    case BLANK:   cell.emptyTag();                                                      break;
                    default:
                }

                // 复制样式
                Integer xf = styleMap.get(cell0.xf);
                if (xf != null) cell.xf = xf;
                else {
                    int style = row0.getCellStyle(cell0);
                    xf = 0;
                    // 字体
                    Font font = styles0.getFont(style);
                    if (font != null) xf |= styles.addFont(font.clone());
                    // 填充
                    Fill fill = styles0.getFill(style);
                    if (fill != null) xf |= styles.addFill(fill.clone());
                    // 边框
                    Border border = styles0.getBorder(style);
                    if (border != null) xf |= styles.addBorder(border.clone());
                    // 格式化
                    NumFmt numFmt = styles0.getNumFmt(style);
                    if (numFmt != null) xf |= styles.addNumFmt(numFmt.clone());
                    // 水平对齐、垂直对齐、自动折行
                    int h = styles0.getHorizontal(style), v = styles0.getVertical(style), w = styles0.getWrapText(style);

                    // 添加进样式表
                    cell.xf = styles.of(xf | h | v | w);
                    styleMap.put(cell0.xf, cell.xf);
                }
            }
        }
    }

    @Override
    public void close() throws IOException {
        super.close();
        // 释放模板流
        if (reader != null) reader.close();
    }
}

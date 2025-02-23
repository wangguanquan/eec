/*
 * Copyright (c) 2017-2023, guanquan.wang@hotmail.com All Rights Reserved.
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
import org.ttzero.excel.entity.e7.XMLWorksheetWriter;
import org.ttzero.excel.entity.style.Border;
import org.ttzero.excel.entity.style.ColorIndex;
import org.ttzero.excel.entity.style.Fill;
import org.ttzero.excel.entity.style.Font;
import org.ttzero.excel.entity.style.NumFmt;
import org.ttzero.excel.entity.style.Styles;
import org.ttzero.excel.manager.Const;
import org.ttzero.excel.util.FileUtil;
import org.ttzero.excel.validation.ListValidation;
import org.ttzero.excel.validation.Validation;
import org.ttzero.excel.reader.Cell;
import org.ttzero.excel.reader.CellType;
import org.ttzero.excel.reader.Col;
import org.ttzero.excel.reader.Dimension;
import org.ttzero.excel.reader.Drawings;
import org.ttzero.excel.reader.ExcelReader;
import org.ttzero.excel.reader.FullSheet;
import org.ttzero.excel.reader.RowSetIterator;
import org.ttzero.excel.util.DateUtil;
import org.ttzero.excel.util.StringUtil;

import java.beans.IntrospectionException;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.AccessibleObject;
import java.lang.reflect.Array;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.nio.ByteBuffer;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.Comparator;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.function.BiFunction;

import static org.ttzero.excel.entity.IWorksheetWriter.isString;
import static org.ttzero.excel.entity.SimpleSheet.defaultDatetimeCell;
import static org.ttzero.excel.entity.style.Styles.INDEX_FONT;
import static org.ttzero.excel.util.ReflectUtil.listDeclaredFieldsUntilJavaPackage;
import static org.ttzero.excel.util.ReflectUtil.readMethodsMap;

/**
 * 模板工作表，它支持指定一个已有的Excel文件作为模板导出，{@code TemplateSheet}将复制模板工作表的样式并替换占位符，
 * 同时{@code TemplateSheet}也可以和其它{@code Worksheet}混用，这意味着可以添加多个模板工作表和普通工作表。
 *
 * <p>创建模板工作表需要指定模板文件，它可以是本地文件也可是输入流{@code InputStream}，支持的类型包含{@code xls}
 * 和{@code xlsx}两种格式，除模板文件外还需要指定工作表，未指定工作表时默认以第一个工作表做为模板。</p>
 *
 * <p>TemplateSheet工作表导出时不受ExcelColumn注解限制，导出的数据范围由模板中的占位符决定，
 * 默认占位符由一对封闭的大括号{@code ${key}}组成，虽然占位符与EL表达式写法相似但模板占位符并不具备EL的能力，
 * 所以无法使用{@code ${1 + 2}}或{@code ${System.getProperty("user.name")}}这类语句来做运算，
 * 占位符<b>仅做替换不做运算</b>所以不需要担心安全漏洞问题。</p>
 *
 * <p>{@link #setData}方法为占位符绑定值，支持对象、Map、Array和List类型，数据量较大时也可以绑定一个数据生产者{@code data-supplier}来分片拉取数据，
 * 它被定义为{@code BiFunction<Integer, T, List<T>>}，其中第一个入参{@code Integer}表示已拉取数据的记录数
 * （并非已写入数据），第二个入参{@code T}表示上一批数据中最后一个对象，业务端可以通过这两个参数来计算下一批数据应该从哪个节点开始拉取，
 * 通常你可以使用第一个参数除以每批拉取的数据大小来确定当前页码，如果数据已排序则可以使用{@code T}对象的排序字段来计算下一批数据的游标以跳过
 * {@code limit ... offset ... }分页查询从而极大提升取数性能。</p>
 *
 * <pre>
 * new Workbook("模板测试")
 *      // 模板工作表
 *     .addSheet(new TemplateSheet(Paths.get("./template.xlsx"))
 *          // 免分页查询用户，根据ID排序并游标拉取
 *         .setData((i,lastOne) -&gt; queryUser(i &gt; 0 ? ((User)lastOne).getId():0))
 *     // 普通对象数组工作表
 *     .addSheet(new ListSheet&lt;&gt;().setData(list))
 *     .writeTo(Paths.get("/tmp/"));</pre>
 *
 * <p>每个占位符都有一个命名空间，使用${namespace.key}这种格式来添加命名空间，默认命名空间为{@code null}。
 * 占位符中还包含三个内置函数它们分别为[&#x40;{@code link:}]、[&#x40;{@code list:}]和[&#x40;{@code media:}]，
 * 分别用于设置单元格的值为超链接、序列和图片，其中序列的值可以从源工作表中获取也可以使用{@link #setData(String, Object)}来设置。
 * <b>注意：内置函数必须独占一个单元格且仅识别固定的三个内置函数，任意其它命令将被识别为普通命名空间</b></p>
 *
 * <p>占位符整体样式：[&#x40;内置函数:][命名空间][.]&lt;占位符&gt;</p>
 *
 * <pre>
 * template.xlsx模板如下：
 * +--------+--------+--------------+---------------+------------------+
 * |  姓名  |  年龄  |     性别     |      头像     |     简历原件     |
 * +--------+--------+--------------+---------------+------------------+
 * |${name} | ${age} | ${&#x40;list:sex} | ${&#x40;media:pic} | ${&#x40;link:jumpUrl} |
 * +--------+--------+--------------+---------------+------------------+
 *
 * // 组装测试数据
 * List&lt;Map&lt;String, Object&gt;&gt; data = new ArrayList&lt;&gt;();
 * Map&lt;String, Object&gt; row1 = new HashMap&lt;&gt;();
 * row1.put("name", "张三");
 * row1.put("age", 26);
 * row1.put("sex", "男");
 * row1.put("pic", Paths.get("./images/head.png"));
 * row1.put("jumpUrl", "https://jianli.com/zhangsan");
 * data.add(row1);
 *
 *  new Workbook("内置函数测试")
 *     // 模板工作表
 *     .addSheet(new TemplateSheet(Paths.get("./template.xlsx"))
 *         // 替换模板中占位符
 *         .setData(data)
 *         // 替换模板中"@list:sex"值为性别序列
 *         .setData("@list:sex", Arrays.asList("未知", "男", "女")))
 *     .writeTo(Paths.get("/tmp/"));</pre>
 *
 * <p>参考文档:</p>
 * <p><a href="https://github.com/wangguanquan/eec/wiki/3-%E6%A8%A1%E6%9D%BF%E5%AF%BC%E5%87%BA">模板导出</a></p>
 *
 * @author guanquan.wang at 2023-12-01 15:10
 */
public class TemplateSheet extends Sheet {
    /**
     * 内置单元格类型-超链接样式
     */
    public static final String HYPERLINK_KEY = "@link:";
    /**
     * 内置单元格类型-图片
     */
    public static final String MEDIA_KEY = "@media:";
    /**
     * 内置单元格类型-序列
     */
    public static final String LIST_KEY = "@list:";
    /**
     * 占位符前缀和后缀
     */
    protected String prefix = "${", suffix = "}";
    /**
     * 模板路径
     */
    protected Path templatePath;
    /**
     * 模板流
     */
    protected InputStream templateStream;
    /**
     * 读取模板用
     */
    protected ExcelReader reader;
    /**
     * 源工作表索引
     */
    protected int originalSheetIndex;
    /**
     * 源工作表名
     */
    protected String originalSheetName;
    /**
     * 行数据迭代器
     */
    protected CommitRowSetIterator rowIterator;
    /**
     * 样式映射，缓存源样式索引映射到目标样式索引
     */
    protected Map<Integer, Integer> styleMap;
    /**
     * 图片
     */
    protected List<Drawings.Picture> pictures;
    /**
     * 以Excel格式输出
     */
    protected boolean writeAsExcel;
    /**
     * 包含占位符的单元格预处理后的结果
     */
    protected PreCell[][] preCells;
    /**
     * 占位符位置标记 pf: 当前占位符的行号 pi: 当前占位符在preNodes的下标
     * afr：auto-filter row
     */
    protected int pf, pi, afr = -1;
    /**
     * 合并单元格（输出时需特殊处理）
     */
    protected List<Dimension> mergeCells;
    /**
     * 缓存源文件合并单元格
     * Key: 首坐标 Value：单元格范围
     */
    protected Map<Long, Dimension> mergeCells0;
    /**
     * 填充数据缓存
     */
    protected Map<String, ValueWrapper> namespaceMapper = new HashMap<>();
    /**
     * 缓存源文件批注
     * Key: 坐标 Value：批注
     */
    protected Map<Long, Comment> comments0;
    /**
     * 实例化模板工作表，默认以第一个工作表做为模板
     *
     * @param templatePath 模板路径
     */
    public TemplateSheet(Path templatePath) {
        this(templatePath, 0);
    }

    /**
     * 实例化模板工作表，默认以第一个工作表做为模板
     *
     * @param name         指定工作表名称
     * @param templatePath 模板路径
     */
    public TemplateSheet(String name, Path templatePath) {
        this(name, templatePath, 0);
    }

    /**
     * 实例化模板工作表并指定模板工作表索引，如果指定索引超过模板Excel中包含的工作表数量则抛异常
     *
     * @param templatePath       模板路径
     * @param originalSheetIndex 指定源工作表索引（从0开始）
     */
    public TemplateSheet(Path templatePath, int originalSheetIndex) {
        this(null, templatePath, originalSheetIndex);
    }

    /**
     * 实例化模板工作表并指定模板工作表索引，如果指定索引超过模板Excel中包含的工作表数量则抛异常
     *
     * @param name               指定工作表名称
     * @param templatePath       模板路径
     * @param originalSheetIndex 指定源工作表索引（从0开始）
     */
    public TemplateSheet(String name, Path templatePath, int originalSheetIndex) {
        this.name = name;
        this.templatePath = templatePath;
        this.originalSheetIndex = originalSheetIndex;
    }

    /**
     * 实例化模板工作表并指定模板工作表名，如果指定源工作表不存在则抛异常
     *
     * @param templatePath      模板路径
     * @param originalSheetName 指定源工作表名
     */
    public TemplateSheet(Path templatePath, String originalSheetName) {
        this(null, templatePath, originalSheetName);
    }

    /**
     * 实例化模板工作表并指定模板工作表名，如果指定源工作表不存在则抛异常
     *
     * @param name              指定工作表名称
     * @param templatePath      模板路径
     * @param originalSheetName 指定源工作表名
     */
    public TemplateSheet(String name, Path templatePath, String originalSheetName) {
        this.name = name;
        this.templatePath = templatePath;
        this.originalSheetName = originalSheetName;
    }

    /**
     * 实例化模板工作表，默认以第一个工作表做为模板
     *
     * @param templateStream 模板输入流
     */
    public TemplateSheet(InputStream templateStream) {
        this(templateStream, 0);
    }

    /**
     * 实例化模板工作表，默认以第一个工作表做为模板
     *
     * @param name           设置工作表名
     * @param templateStream 模板输入流
     */
    public TemplateSheet(String name, InputStream templateStream) {
        this(name, templateStream, 0);
    }

    /**
     * 实例化模板工作表并指定模板工作表索引，如果指定索引超过模板Excel中包含的工作表数量则抛异常
     *
     * @param templateStream     模板输入流
     * @param originalSheetIndex 指定源工作表索引
     */
    public TemplateSheet(InputStream templateStream, int originalSheetIndex) {
        this(null, templateStream, originalSheetIndex);
    }

    /**
     * 实例化模板工作表并指定模板工作表名，如果指定源工作表不存在则抛异常
     *
     * @param templateStream    模板输入流
     * @param originalSheetName 指定源工作表名
     */
    public TemplateSheet(InputStream templateStream, String originalSheetName) {
        this(null, templateStream, originalSheetName);
    }

    /**
     * 实例化模板工作表并指定模板工作表索引，如果指定索引超过模板Excel中包含的工作表数量则抛异常
     *
     * @param name               设置工作表名
     * @param templateStream     模板输入流
     * @param originalSheetIndex 指定源工作表索引
     */
    public TemplateSheet(String name, InputStream templateStream, int originalSheetIndex) {
        this.name = name;
        this.templateStream = templateStream;
        this.originalSheetIndex = originalSheetIndex;
    }

    /**
     * 实例化模板工作表并指定模板工作表名，如果指定源工作表不存在则抛异常
     *
     * @param name              设置工作表名
     * @param templateStream    模板输入流
     * @param originalSheetName 指定源工作表名
     */
    public TemplateSheet(String name, InputStream templateStream, String originalSheetName) {
        this.name = name;
        this.templateStream = templateStream;
        this.originalSheetName = originalSheetName;
    }

    /**
     * 设置占位符前缀，默认前缀为{@code $&#x123;}
     *
     * @param prefix 占位符前缀
     * @return 当前工作表
     */
    public TemplateSheet setPrefix(String prefix) {
        if (StringUtil.isBlank(prefix))
            throw new IllegalArgumentException("Illegal prefix value");
        this.prefix = prefix;
        return this;
    }

    /**
     * 设置占位符后缀，默认后缀为{@code &#x125;}
     *
     * @param suffix 占位符后缀
     * @return 当前工作表
     */
    public TemplateSheet setSuffix(String suffix) {
        if (StringUtil.isBlank(suffix))
            throw new IllegalArgumentException("Illegal suffix value");
        this.suffix = suffix;
        return this;
    }

    /**
     * 绑定数据到默认命名空间，默认命名空间为{@code null}
     *
     * @param o 任意对象，可以为Java Bean，Map，或者数组
     * @return 当前工作表
     */
    public TemplateSheet setData(Object o) {
        return setData(null, o);
    }

    /**
     * 绑定数据到指定命名空间上
     *
     * @param namespace 命名空间
     * @param o         任意对象，可以为Java Bean，Map，或者数组
     * @return 当前工作表
     */
    public TemplateSheet setData(String namespace, Object o) {
        if ("this".equals(namespace)) namespace = null;
        ValueWrapper vw = namespaceMapper.get(namespace);
        if (vw == null) {
            vw = new ValueWrapper();
            namespaceMapper.put(namespace, vw);
        } else LOGGER.warn("The namespace[{}] already exists.", namespace);
        if (o == null) vw.option = 0;
        else {
            Class<?> clazz = o.getClass();
            if (Map.class.isAssignableFrom(clazz)) {
                vw.option = 2;
                Map map = (Map) o;
                if (vw.map == null) vw.map = map;
                else vw.map.putAll(map);
            }
            else if (List.class.isAssignableFrom(clazz)) {
                List list = (List) o;
                Object oo = getFirstObject(list);
                if (oo != null) {
                    vw.option = Map.class.isAssignableFrom(oo.getClass()) ? 3 : 4;
                    if (vw.list == null) vw.list = list; else vw.list.addAll(list);
                    if (vw.option == 4) vw.accessibleObjectMap = parseClass(oo.getClass());
                } else vw.option = 0;
            }
            else if (clazz.isArray()) {
                int len = Array.getLength(o);
                if (vw.list == null) vw.list = new ArrayList<>(len);
                for (int i = 0; i < len; i++) {
                    Object oo = Array.get(o, i);
                    vw.list.add(oo);
                    if (oo != null && vw.option == 0) {
                        vw.option = Map.class.isAssignableFrom(oo.getClass()) ? 3 : 4;
                        if (vw.option == 4) vw.accessibleObjectMap = parseClass(oo.getClass());
                    }
                }
            }
            else {
                vw.o = o;
                vw.option = 1;
                vw.accessibleObjectMap = parseClass(clazz);
            }
        }
        return this;
    }

    /**
     * 绑定一个{@code Supplier}到默认命名空间，适用于未知长度或数量最大的数组
     *
     * @param dataSupplier 数据产生者
     * @return 当前工作表
     */
    public TemplateSheet setData(BiFunction<Integer, Object, List<?>> dataSupplier) {
        return setData(null, dataSupplier);
    }

    /**
     * 绑定一个{@code Supplier}到指定命名空间，适用于未知长度或数量最大的数组
     *
     * @param namespace 命名空间
     * @param dataSupplier  数据产生者
     * @return 当前工作表
     */
    public TemplateSheet setData(String namespace, BiFunction<Integer, Object, List<?>> dataSupplier) {
        if ("this".equals(namespace)) namespace = null;
        ValueWrapper vw = namespaceMapper.get(namespace);
        if (vw != null) {
            LOGGER.warn("The namespace[{}] already exists.", namespace);
        } else {
            vw = new ValueWrapper();
            namespaceMapper.put(namespace, vw);
        }
        vw.supplier = dataSupplier;

        // 加载第一批数据预处理数据类型
        if (dataSupplier != null) {
            List list = dataSupplier.apply(0, null);
            Object oo = getFirstObject(list);
            if (oo != null) {
                vw.size += list.size();
                if (vw.list == null) vw.list = list;
                else vw.list.addAll(list);
                vw.option = Map.class.isAssignableFrom(oo.getClass()) ? 3 : 4;
                if (vw.option == 4) vw.accessibleObjectMap = parseClass(oo.getClass());
            } else vw.option = 0;
        }
        return this;
    }

    /**
     * 获取下一段{@link RowBlock}行块数据，工作表输出协议通过此方法循环获取行数据并落盘，
     * 行块被设计为一个滑行窗口，下游输出协议只能获取一个窗口的数据默认包含32行。
     *
     * @return 行块
     */
    public RowBlock nextBlock() {
        rowBlock.clear();
        // 装载数据（这里不需要判断是否有表头，模板不需要表头）
        resetBlockData();
        return rowBlock.flip();
    }

    @Override
    public Column[] getAndSortHeaderColumns() {
        if (!headerReady) {
            // 解析模板工作表并复制信息到当前工作表中
            int size;
            try {
                size = init();
            } catch (IOException e) {
                throw new ExcelWriteException(e);
            }
            if (size <= 0) columns = new Column[0];
            else {
                sortColumns(columns);
                calculateRealColIndex();
                resetCommonProperties(columns);
            }
            markExtProp();
            headerReady = true;
        }
        return columns;
    }

    /**
     * 读取模板头信息并复杂到当前工作表
     *
     * @return 列的个数
     * @throws IOException 读取模板异常
     */
    protected int init() throws IOException {
        // 实例化ExcelReader
        if (templatePath != null) reader = ExcelReader.read(templatePath);
        else if (templateStream != null) reader = ExcelReader.read(templateStream);

        // 查找源工作表
        org.ttzero.excel.reader.Sheet[] sheets = reader.all();
        if (StringUtil.isNotBlank(originalSheetName)) {
            int index = 0;
            for (; index < sheets.length && !originalSheetName.equals(sheets[index].getName()); index++) ;
            if (index >= sheets.length)
                throw new IOException("The original worksheet [" + originalSheetName + "] does not exist in template file.");
            originalSheetIndex = index;
        } else if (originalSheetIndex < 0 || originalSheetIndex >= sheets.length)
            throw new IOException("The original worksheet index [" + originalSheetIndex + "] is out of range in template file[0-" + sheets.length + "].");

        // 加载模板工作表
        FullSheet sheet = reader.sheet(originalSheetIndex).asFullSheet();
        writeAsExcel = sheetWriter != null && XMLWorksheetWriter.class.isAssignableFrom(sheetWriter.getClass());

        // 解析公共信息
        int n = prepareCommonData(sheet);
        // 预处理样式和占位符
        rowIterator = prepare(sheet);
        pf = preCells == null ? -1 : preCells[0][0].row;

        // 忽略表头输出
        super.ignoreHeader();

        // 解析公共信息
        return n;
    }

    @Override
    protected void resetBlockData() {
        Dimension mergeCell;
        PreCell pn;
        Comment comment;
        Column emptyColumn = new Column();
        for (int rbs = rowBlock.capacity(), n = 0, limit = sheetWriter.getRowLimit(), len; n++ < rbs && rows < limit && rowIterator.hasNext(); ) {
            Row row = rowBlock.next();
            org.ttzero.excel.reader.Row row0 = rowIterator.next();
            row.index = rows = rowIterator.rows - 1;
            row.height = row0.getHeight();
            row.hidden = row0.isHidden();
            // 空行特殊处理（lc-fc=-1)
            len = Math.max(row0.getLastColumnIndex(), 0);
            Cell[] cells = row.realloc(len);
            // 预处理
            if (row0.getRowNum() == pf && !rowIterator.hasFillCell) rowIterator.withPreNodes(preCells[pi], namespaceMapper);

            for (int i = 0; i < len; i++) {
                Cell cell = cells[i], cell0 = row0.getCell(i);
                // Clear cells
                cell.clear();

                // 复制样式
                cell.xf = styleMap.getOrDefault(cell0.xf, 0);
                if (cell.h) cell.xf = hyperlinkStyle(workbook.getStyles(), cell.xf);

                boolean fillCell = false;
                // 复制数据
                switch (row0.getCellType(cell0)) {
                    case STRING:
                        if (rowIterator.hasFillCell && (pn = rowIterator.preNodes[i]) != null) {
                            fillCell = true;
                            fillValue(row, cell, pn, emptyColumn);

                            // 处理单行合并单元格
                            if (pn.m != null) {
                                // 正数为行合并 负数为列合并
                                if (pn.m > 0) mergeCells.add(new Dimension(rows + 1, (short) (i + 1), rows + 1, (short) (i + pn.m + 1)));
                                else mergeCells.add(new Dimension(rows + 1, (short) (i + 1), rows + ~pn.m, (short) (i + 1)));
                            }
                        } else cell.setString(row0.getString(cell0));
                        break;
                        // FIXME 范围外的数据不需要复制，要继续向后走
                    case LONG:    cell.setLong(row0.getLong(cell0));                                    break;
                    case INTEGER: cell.setInt(row0.getInt(cell0));                                      break;
                    case DECIMAL: cell.setDecimal(row0.getDecimal(cell0));                              break;
                    case DOUBLE:  cell.setDouble(row0.getDouble(cell0));                                break;
                    case DATE:    cell.setDateTime(DateUtil.toDateTimeValue(row0.getTimestamp(cell0))); break;
                    case BOOLEAN: cell.setBool(row0.getBoolean(cell0));                                 break;
                    case BLANK:   cell.emptyTag();                                                      break;
                    default:
                }

                if (!writeAsExcel) continue;

                // TODO 复制公式（不是简单的复制，需重新计算位置）
                if (row0.hasFormula(cell0)) cell.setFormula(row0.getFormula(cell0));

                long k = dimensionKey(row0.getRowNum() - 1, i);
                // 合并单元格重新计算位置
                if (!fillCell && mergeCells0 != null && (mergeCell = mergeCells0.get(k)) != null) {
                    if (rows <= row0.getRowNum()) mergeCells.add(mergeCell);
                    else {
                        int r = rows - row0.getRowNum() + 1;
                        mergeCells.add(new Dimension(mergeCell.firstRow + r, mergeCell.firstColumn, mergeCell.lastRow + r, mergeCell.lastColumn));
                    }
                }
                if (comments0 != null && (comment = comments0.get(k)) != null) {
                    createComments().addComment(rows + 1, i + 1, comment);
                }
            }

            // 写入一行数据末尾处理
            rowEnd(row0, row);
        }
    }

    protected void rowEnd(org.ttzero.excel.reader.Row row0, Row row) {
        // 占位符是否已消费结束
        boolean consumerEnd = true;
        if (!rowIterator.consumerNamespaces.isEmpty()) {
            for (String vwKey : rowIterator.consumerNamespaces) {
                ValueWrapper vw = namespaceMapper.get(vwKey);
                if (++vw.i < vw.list.size()) consumerEnd = false;
                    // 加载更多数据
                else if (vw.supplier != null) {
                    List list = vw.supplier.apply(vw.size, !vw.list.isEmpty() ? vw.list.get(vw.list.size() - 1) : null);
                    if (list != null && !list.isEmpty()) {
                        vw.list = list;
                        vw.i = 0;
                        vw.size += list.size();
                        consumerEnd = false;
                    } else vw.option = -1;
                } else vw.option = -1;
            }
        }
        // Rrk
        if (consumerEnd) rowCommit(row0, row);
    }

    protected void rowCommit(org.ttzero.excel.reader.Row row0, org.ttzero.excel.entity.Row row) {
        PreCell pn;
        Object e;
        int len = Math.max(row0.getLastColumnIndex(), 0);
        if (rowIterator.hasFillCell) {
            for (int i = row0.getFirstColumnIndex(); i < len; i++) {
                if ((pn = rowIterator.preNodes[i]) != null && pn.validation != null) {
                    Dimension sqref = pn.validation.sqref;
                    pn.validation.sqref = new Dimension(sqref.firstRow, sqref.firstColumn, sqref.lastRow + pn.v, sqref.firstColumn);
                }
            }
            pi++;
            pf = preCells.length > pi && preCells[pi] != null && preCells[pi].length >= 1 ? preCells[pi][0].row : -1;
        }
        // 过滤行列重算
        if (afr == row0.getRowNum() && (e = getExtPropValue(Const.ExtendPropertyKey.AUTO_FILTER)) instanceof Dimension) {
            Dimension autoFilter = (Dimension) e;
            putExtProp(Const.ExtendPropertyKey.AUTO_FILTER, new Dimension(autoFilter.firstRow + row.getIndex() - afr + 1, autoFilter.firstColumn, autoFilter.getLastRow() + row.getIndex() - afr + 1, autoFilter.lastColumn));
            afr = -1;
        }
        rowIterator.commit();
    }

    /**
     * 获取占位符的实际值
     *
     * @param node 占位符节点信息
     * @return 值
     */
    protected Object getNodeValue(Node node) {
        // 纯文本
        if ((node.option & 1) == 0) return node.val;
        ValueWrapper vw = namespaceMapper.get(node.namespace);
        Object e = null;
        if (vw != null) {
            switch (vw.option) {
                case 1: e = getObjectValue(vw.accessibleObjectMap.get(node.val), vw.o, LOGGER, node.val);              break;
                case 2: e = vw.map.get(node.val);                                                                      break;
                case 3: e = ((Map<String, Object>) vw.list.get(vw.i)).get(node.val);                                   break;
                case 4: e = getObjectValue(vw.accessibleObjectMap.get(node.val), vw.list.get(vw.i), LOGGER, node.val); break;
                default:
            }
        }
        return e;
    }

    /**
     * 反射获取对象的值
     *
     * @param ao     Method 或 Field
     * @param o      对象
     * @param logger 日志
     * @param key    占位符
     * @return 值
     */
    protected static Object getObjectValue(AccessibleObject ao, Object o, Logger logger, String key) {
        Object e = null;
        try {
            if (ao instanceof Method) e = ((Method) ao).invoke(o);
            else if (ao instanceof Field) e = ((Field) ao).get(o);
            else e = o;
        } catch (IllegalAccessException | InvocationTargetException ex) {
            logger.warn("Invoke " + key + " value error", ex);
        }
        return e;
    }

    protected void fillValue(Row row, Cell cell, PreCell pn, Column emptyColumn) {
        Object e;
        if (pn.nodes.length == 1) {
            e = getNodeValue(pn.nodes[0]);
            int type = pn.nodes[0].getType();
            if (e != null) {
                Class<?> clazz = e.getClass();
                switch (type) {
                    // Hyperlink
                    case 1: if (isString(clazz)) cell.setHyperlink(e.toString()); break;
                    // Media
                    case 2:
                        if (isString(clazz)) {
                            cellValueAndStyle.writeAsMedia(row, cell, e.toString(), emptyColumn, String.class);
                        } else if (Path.class.isAssignableFrom(clazz)) {
                            cell.setPath((Path) e);
                        } else if (File.class.isAssignableFrom(clazz)) {
                            cell.setPath(((File) e).toPath());
                        } else if (InputStream.class.isAssignableFrom(clazz)) {
                            cell.setInputStream((InputStream) e);
                        } else if (clazz == byte[].class) {
                            cell.setBinary((byte[]) e);
                        } else if (ByteBuffer.class.isAssignableFrom(clazz)) {
                            cell.setByteBuffer((ByteBuffer) e);
                        }
                        break;
                    default:
                        emptyColumn.setClazz(clazz);
                        cellValueAndStyle.setCellValue(row, cell, e, emptyColumn, clazz, false);
                        // 日期类型添加默认format
                        if (cell.t == Cell.DATETIME || cell.t == Cell.DATE || cell.t == Cell.TIME) {
                            datetimeCell(workbook.getStyles(), cell);
                        }
                }
            } else cell.emptyTag();
            // 序列的dimension纵向+1
            if (type == 3) pn.v++;
        } else {
            int k = 0;
            for (Node node : pn.nodes) {
                e = getNodeValue(node);
                if (e != null) {
                    String s = e.toString();
                    int vn = s.length();
                    if (vn + k > pn.cb.length) pn.cb = Arrays.copyOf(pn.cb, vn + k + 128);
                    s.getChars(0, vn, pn.cb, k);
                    k += vn;
                }
            }
            cell.setString(new String(pn.cb, 0, k));
        }
    }

    /**
     * 日期类型添加默认format
     *
     * @param styles Styles
     * @param cell 单元格
     */
    protected void datetimeCell(Styles styles, Cell cell) {
        defaultDatetimeCell(styles, cell);
    }

    @Override
    public void afterSheetDataWriter(int total) {
        super.afterSheetDataWriter(total);

        // 添加图片
        if (pictures != null) {
            try {
                for (Drawings.Picture p : pictures) {
                    if (p == null || p.isBackground()) continue;
                    sheetWriter.writePicture(toWritablePicture(p));
                }
            } catch (IOException e) {
                LOGGER.warn("Copy pictures failed.", e);
            }
        }

        // 添加合并
        if (mergeCells != null) putExtProp(Const.ExtendPropertyKey.MERGE_CELLS, mergeCells);
        // TODO 重置过滤位置
        Object o = getExtPropValue(Const.ExtendPropertyKey.AUTO_FILTER);
        if (o instanceof Dimension) {
            Dimension autoFilter = (Dimension) o;
            putExtProp(Const.ExtendPropertyKey.AUTO_FILTER, autoFilter);
        }
    }

    @Override
    public void close() throws IOException {
        super.close();
        if (reader != null) {
            reader.close();
            reader = null;
        }
        if (templateStream != null) {
            templateStream.close();
            templateStream = null;
        }
        rowIterator = null;
        namespaceMapper = null;
        columns = null;
        relManager = null;
        rowBlock = null;
        sheetWriter = null;
        cellValueAndStyle = null;
        extProp = null;
    }

    /**
     * 将图片转为导出格式
     *
     * @param pic 图片
     * @return 可导出图片
     */
    public static Picture toWritablePicture(Drawings.Picture pic) {
        Picture p = new Picture();
        p.localPath = pic.getLocalPath();
        p.row = pic.getDimension().firstRow - 1;
        p.col = pic.getDimension().firstColumn - 1;
        p.toRow = pic.getDimension().lastRow - 1;
        p.toCol = pic.getDimension().lastColumn - 1;
        p.padding = pic.getPadding();
        p.revolve = pic.getRevolve();
        p.property = pic.getProperty();
        p.effect = pic.getEffect();
        return p;
    }

    /**
     * 预处理样式和占位符
     *
     * @param originalSheet 模板工作表
     * @return 模板工作表行迭代器
     */
    protected CommitRowSetIterator prepare(org.ttzero.excel.reader.Sheet originalSheet) {
        // 模板文件样式
        Styles styles0 = reader.getStyles(), styles = workbook.getStyles();
        // 样式缓存
        if (styleMap == null) styleMap = writeAsExcel ? new HashMap<>() : Collections.emptyMap();
        int prefixLen = prefix.length(), suffixLen = suffix.length(), pf = 0;
        for (Iterator<org.ttzero.excel.reader.Row> iter = originalSheet.iterator(); iter.hasNext(); ) {
            org.ttzero.excel.reader.Row row = iter.next();
            int index = 0;
            for (int i = row.getFirstColumnIndex(), end = row.getLastColumnIndex(); i < end; i++) {
                Cell cell = row.getCell(i);

                // 复制样式
                if (writeAsExcel && !styleMap.containsKey(cell.xf)) {
                    // 复制样式添加进样式表
                    styleMap.put(cell.xf, copyStyle(styles0, styles, cell.xf));
                }

                // 判断字符串是否包含占位符，可以是一个或多个
                if (row.getCellType(cell) == CellType.STRING) {
                    String v = row.getString(cell);
                    // 预处理单元格的值
                    PreCell preCell = prepareCellValue(row.getRowNum(), i, v, prefixLen, suffixLen);
                    if (preCell != null) {
                        if (preCells == null) preCells = new PreCell[10][];
                        PreCell[] pns;
                        if (index == 0) {
                            if (pf >= preCells.length) preCells = Arrays.copyOf(preCells, preCells.length + 10);
                            preCells[pf++] = pns = new PreCell[Math.min(end - i, 10)];
                        } else if (index >= (pns = preCells[pf - 1]).length)
                            preCells[pf - 1] = pns = Arrays.copyOf(pns, pns.length + 10);
                        pns[index++] = preCell;
                    }
                }
            }
            if (index > 0 && preCells[pf - 1].length > index) preCells[pf - 1] = Arrays.copyOf(preCells[pf - 1], index);
        }

        return new CommitRowSetIterator((RowSetIterator) originalSheet.reset().iterator());
    }

    /**
     * 复制样式
     *
     * @param srcStyles 源样式表
     * @param distStyle 目标样式表
     * @param copyXf    样式索引
     * @return 复制样式在目标样式表中的索引
     */
    protected static int copyStyle(Styles srcStyles, Styles distStyle, int copyXf) {
        int style =  srcStyles.getStyleByIndex(copyXf), xf = 0;
        // 字体
        Font font = srcStyles.getFont(style);
        if (font != null) xf |= distStyle.addFont(font.clone());
        // 填充
        Fill fill = srcStyles.getFill(style);
        if (fill != null) xf |= distStyle.addFill(fill.clone());
        // 边框
        Border border = srcStyles.getBorder(style);
        if (border != null) xf |= distStyle.addBorder(border.clone());
        // 格式化
        NumFmt numFmt = srcStyles.getNumFmt(style);
        if (numFmt != null) xf |= distStyle.addNumFmt(numFmt.clone());
        // 水平对齐
        xf |= srcStyles.getHorizontal(style);
        // 垂直对齐
        xf |= srcStyles.getVertical(style);
        // 自动折行
        xf |= srcStyles.getWrapText(style);
        return distStyle.of(xf);
    }

    /**
     * 解析公共数据
     *
     * @param originalSheet 源模板工作表
     * @return 列数
     */
    protected int prepareCommonData(org.ttzero.excel.reader.FullSheet originalSheet) {
        // 获取列属性
        int len = 0;
        List<Col> cols = originalSheet.getCols();
        if (cols != null && !cols.isEmpty()) {
            cols.sort(Comparator.comparingInt(a -> a.max));
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
                    if (col.styleIndex > 0) c.globalStyleIndex = col.styleIndex;
                    columns[i++] = c;
                }
            }
        }

        // xlsx格式输出才进行以下格式复制
        if (!writeAsExcel) return len;

        // 冻结,直接复制不需要计算移动
        Panes panes = originalSheet.getFreezePanes();
        if (panes != null) putExtProp(Const.ExtendPropertyKey.FREEZE, panes);

        // 合并
        List<Dimension> mergeCells0 = originalSheet.getMergeCells();
        if (mergeCells0 != null) {
            mergeCells = new ArrayList<>(mergeCells0.size());
            this.mergeCells0 = new HashMap<>(mergeCells0.size());
            // 这里将坐标切换到 base 0，方便后续取值
            for (Dimension dim : mergeCells0) this.mergeCells0.put(dimensionKey(dim.firstRow - 1, dim.firstColumn - 1), dim);
        }

        // 过滤
        Dimension autoFilter = originalSheet.getFilter();
        if (autoFilter != null) {
            afr = autoFilter.getFirstRow();
            putExtProp(Const.ExtendPropertyKey.AUTO_FILTER, autoFilter);
        }

        // 是否显示网格线
        this.showGridLines = originalSheet.isShowGridLines();

        // 是否隐藏
        this.hidden = originalSheet.isHidden();

        // 预置列宽
        double defaultColWidth = originalSheet.getDefaultColWidth(), defaultRowHeight = originalSheet.getDefaultRowHeight();
        if (defaultColWidth >= 0) putExtProp("defaultColWidth", defaultColWidth);
        if (defaultRowHeight >= 0) putExtProp("defaultRowHeight", defaultRowHeight);

        // 是否有缩放
        Integer zoomScale = originalSheet.getZoomScale();
        if (zoomScale != null) putExtProp(Const.ExtendPropertyKey.ZOOM_SCALE, zoomScale);

        // FIXME 图片（较为复杂不能简单复制，需要计算中间插入或扣除的行）
        try {
            List<Drawings.Picture> pictures = originalSheet.listPictures();
            if (pictures != null && !pictures.isEmpty()) {
                this.pictures = pictures.size() > 1 || !pictures.get(0).isBackground() ? new ArrayList<>(pictures.size()) : null;
                for (Drawings.Picture p : pictures) {
                    if (FileUtil.exists(p.getLocalPath())) {
                        if (p.isBackground()) setWaterMark(WaterMark.of(p.getLocalPath()));
                        else this.pictures.add(p);
                    }
                }
            }
        } catch (Exception ex) {
            // Ignore
        }

        // 批注
        Map<Long, Comment> commentMap = originalSheet.getComments();
        if (!commentMap.isEmpty()) {
            this.comments0 = new HashMap<>(commentMap.size());
            for (Map.Entry<Long, Comment> entry : commentMap.entrySet()) {
                this.comments0.put(dimensionKey(((int) (entry.getKey() >> 16)) - 1, ((int) (entry.getKey() & 65535)) - 1), entry.getValue());
            }
        }

        // 样式缓存
        if (styleMap == null) styleMap = new HashMap<>();
        // 复制全局样式
        if (columns != null) {
            Styles styles0 = reader.getStyles(), styles = workbook.getStyles();
            for (Column col : columns) {
                // 存在全局列样式添加进样式表
                if (col.globalStyleIndex > 0) {
                    int newXf = styleMap.getOrDefault(col.globalStyleIndex, -1);
                    if (newXf == -1) {
                        newXf = copyStyle(styles0, styles, col.globalStyleIndex);
                        styleMap.put(col.globalStyleIndex, newXf);
                    }
                    col.globalStyleIndex = newXf;
                }
            }
        }

        return len;
    }

    /**
     * 单元格字符串预处理，检测是否包含占位符以及占位符预处理
     *
     * @param row       单元格行坐标(one base)
     * @param col       单元格列坐标(zero base)
     * @param v         单元格原始值
     * @param prefixLen 占位符前缀长度
     * @param suffixLen 占位符后缀长度
     * @return 预处理结果
     */
    protected PreCell prepareCellValue(int row, int col, String v, int prefixLen, int suffixLen) {
        int len = v.length();
        if (len <= prefixLen + suffixLen) return null;
        int pi = 0, fi = v.indexOf(prefix);
        if (fi < 0) return null;
        int fn, li = v.indexOf(suffix, fn = fi + prefixLen);
        if (li <= fn) return null;

        PreCell pn = new PreCell();
        pn.row = row; pn.col = col;
        do {
            if (fi > pi) {
                Node node = new Node();
                if (pn.nodes == null) pn.nodes = new Node[] { node };
                else {
                    pn.nodes = Arrays.copyOf(pn.nodes, pn.nodes.length + 1);
                    pn.nodes[pn.nodes.length - 1] = node;
                }
                node.val = v.substring(pi, fi);
            }

            int j = fn;
            for (; j < li && v.charAt(j) != '.'; j++) ;

            Node node = new Node();
            node.option = 1;
            if (pn.nodes == null) pn.nodes = new Node[] { node };
            else {
                pn.nodes = Arrays.copyOf(pn.nodes, pn.nodes.length + 1);
                pn.nodes[pn.nodes.length - 1] = node;
            }
            // 包含namespace，可能为数组或多级对象
            if (j < li) {
                node.namespace = v.substring(fn, j).trim();
                node.val = v.substring(j + 1, li).trim();
            } else node.val = v.substring(fn, li).trim();

            pi = li + suffixLen;
            fi = v.indexOf(prefix, pi);
            if (fi < 0) break;
            li = v.indexOf(suffix, fn = fi + prefixLen);
            if (li <= fn) break;
        } while (fi < len);

        // 尾部字符
        if (pi < len) {
            Node node = new Node();
            pn.nodes = Arrays.copyOf(pn.nodes, pn.nodes.length + 1);
            pn.nodes[pn.nodes.length - 1] = node;
            node.val = v.substring(pi);
        }

        // 分配临时空间后续共享使用
        if (pn.nodes.length > 1) pn.cb = new char[Math.max(len + (len >> 1), 128)];
        return afterParseCell(pn);
    }

    /**
     * 解析完占位符后调用此方法，可以处理内置函数
     *
     * @param pn 原始占位符预处理
     * @return 占位符预处理
     */
    protected PreCell afterParseCell(PreCell pn) {
        // 检测是否包含单行合并单元格
        if (mergeCells0 != null) {
            Dimension mc = mergeCells0.get(dimensionKey(pn.row - 1, pn.col));
            // FIXME 目前只支持单行合并
            if (mc != null && mc.height == 1) pn.m = mc.width - 1;
        }

        // 处理内置函数
        if (pn.nodes.length == 1 && (pn.nodes[0].option & 1) == 1) {
            Node node = pn.nodes[0];
            boolean useNamespace = StringUtil.isNotEmpty(node.namespace);
            String k = useNamespace ? node.namespace : node.val;
            if (k.length() < 7 || k.charAt(0) != '@') return pn;
            String[] keys = {HYPERLINK_KEY, MEDIA_KEY, LIST_KEY};
            int p = prefixMatch(keys, k);
            if (p >= 0) {
                node.option |= (p + 1) << 1;
                pn.v = 0;
                String innerFormulaStr = useNamespace ? node.namespace + '.' + node.val : node.val;
                int pLen = keys[p].length();
                if (useNamespace) node.namespace = node.namespace.substring(pLen);
                else node.val = node.val.substring(pLen);
                ValueWrapper vw = namespaceMapper.get(innerFormulaStr);
                if (vw != null && vw.option == 4) {
                    // TODO 读取源文件中的数据验证
                    pn.validation = new ListValidation<>().in(vw.list).dimension(new Dimension(pn.row, (short) (pn.col + 1)));
                    Object o = getExtPropValue(Const.ExtendPropertyKey.DATA_VALIDATION);
                    List<Validation> validations;
                    if (o instanceof List) validations = (List) o;
                    else putExtProp(Const.ExtendPropertyKey.DATA_VALIDATION, validations = new ArrayList<>());
                    // 数据校验
                    validations.add(pn.validation);
                }
            }
        }
        return pn;
    }

    protected int[] fontIndices = {-1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1};
    protected int hyperlinkStyle(Styles styles, int xf) {
        int style = styles.getStyleByIndex(xf);
        int fontIndex = Math.max(0, style << 8 >>> (INDEX_FONT + 8)), fi;
        if (fontIndex > fontIndices.length) {
            int n = fontIndices.length;
            fontIndices = Arrays.copyOf(fontIndices, Math.min(n + 16, fontIndex));
            Arrays.fill(fontIndices, n, fontIndices.length, -1);
        }
        if ((fi = fontIndices[fontIndex]) == -1) {
            Font font = styles.getFont(style).clone();
            font.setStyle(0).underline();
            font.setColor(ColorIndex.themeColors[10]);
            fontIndices[fontIndex] = fi = workbook.getStyles().addFont(font);
        }
        return workbook.getStyles().of(Styles.clearFont(style) | fi);
    }

    /**
     * 解析对象的方法和字段，方便后续取数
     *
     * @param clazz 待解析的对象
     * @return Key：字段名 Value: Method/Field
     */
    protected Map<String, AccessibleObject> parseClass(Class<?> clazz) {
        Map<String, AccessibleObject> tmp = new HashMap<>();
        try {
            tmp.putAll(readMethodsMap(clazz, Object.class));

            Field[] declaredFields = listDeclaredFieldsUntilJavaPackage(clazz);
            for (Field f : declaredFields) {
                if (!tmp.containsKey(f.getName())) {
                    f.setAccessible(true);
                    tmp.put(f.getName(), f);
                }
            }
        } catch (IntrospectionException e) {
            LOGGER.warn("Get class {} methods failed.", clazz);
        }
        return tmp;
    }

    /**
     * 首行首列进行计算后转为缓存的Key
     *
     * @param row 行号(zero base)
     * @param col 列号(zero base)
     * @return 缓存Key
     */
    public static long dimensionKey(int row, int col) {
        return col & 0x7FFF | ((long) row) << 16;
    }

    /**
     * 获取数组中第一个非{@code null}值
     *
     * @param list 数组
     * @return 第一个非{@code null}值
     */
    protected static Object getFirstObject(List<?> list) {
        if (list == null || list.isEmpty()) return null;
        Object first = list.get(0);
        if (first != null) return first;
        int i = 1, len = list.size();
        do {
            first = list.get(i++);
        } while (first == null && i < len);
        return first;
    }

    /**
     * 前缀匹配，数组中的词为前缀词
     *
     * @param array 前缀词
     * @param v     匹配词
     * @return 匹配词在前缀词
     */
    public static int prefixMatch(String[] array, String v) {
        int i = 0, len = array.length;
        for (; i < len && !v.startsWith(array[i]); i++) ;
        return i < len ? i : -1;
    }

    /**
     * 需要手动调用{@link #commit()}方法才会移动游标
     */
    public static class CommitRowSetIterator implements Iterator<org.ttzero.excel.reader.Row> {
        public final RowSetIterator iterator;
        public org.ttzero.excel.reader.Row current;
        public int rows;
        public PreCell[] preNodes;
        public boolean hasFillCell;
        public Set<String> consumerNamespaces = new HashSet<>();

        public CommitRowSetIterator(RowSetIterator iterator) {
            this.iterator = iterator;
        }

        @Override
        public boolean hasNext() {
            return current != null || iterator.hasNext();
        }

        @Override
        public org.ttzero.excel.reader.Row next() {
            org.ttzero.excel.reader.Row row = current != null ? current : (current = iterator.next());
            rows = Math.max(row.getRowNum(), rows + 1);
            return row;
        }

        public void withPreNodes(PreCell[] preNodes, Map<String, ValueWrapper> namespaceMapper) {
            int len = current.getLastColumnIndex(), len0 = preNodes[preNodes.length - 1].col + 1;
            this.preNodes = new PreCell[Math.max(len, len0)];
            for (PreCell p : preNodes) this.preNodes[p.col] = p;
            hasFillCell = true;

            if (!namespaceMapper.isEmpty()) {
                ValueWrapper vw;
                for (PreCell pn : preNodes) {
                    for (Node node : pn.nodes) {
                        if ((node.option & 1) == 1 && (vw = namespaceMapper.get(node.namespace)) != null && (vw.option == 3 || vw.option == 4))
                            consumerNamespaces.add(node.namespace);
                    }
                }
            }
        }

        /**
         * 提交后才将移动到下一行，否则一直停留在当前行
         */
        public void commit() {
            current = null;
            preNodes = null;
            hasFillCell = false;
            consumerNamespaces.clear();
        }
    }

    /**
     * 预处理单元格
     */
    public static class PreCell {
        /**
         * 单元格行列值，行从1开始 列从0开始
         */
        public int row, col;
        /**
         * 节点信息
         */
        public Node[] nodes;
        /**
         * 共享空间
         */
        public char[] cb;
        /**
         * m: 合并范围 正数为行合并 负数为列合并
         * v: 数据验证范围
         */
        public Integer m, v;
        /**
         * 数据验证
         */
        public Validation validation;
    }

    /**
     * 单元格预处理节点
     */
    public static class Node {
        /**
         * 标志位集合，保存一些简单的标志位以节省空间，对应的位点说明如下
         *
         * <blockquote><pre>
         *  Bit | Contents
         * --- -+---------
         * 7, 1 | 是否为占位符 0: 普通文本 1: 占位符
         * 6, 3 | 1: 超链接 2: 图片 3: 序列
         * </pre></blockquote>
         */
        public byte option;
        /**
         * 命令空间
         */
        public String namespace;
        /**
         * 原始文本或占位符
         */
        public String val;

        /**
         * 返回单元格的值类型
         *
         * @return 0: 文本 1: 超链接 2: 图片 3: 序列
         */
        public int getType() {
            return (option >>> 1) & 7;
        }

        @Override
        public String toString() {
            return (option & 1) == 0 ? val : ('$' + (namespace != null ? namespace + "::" : "") + val);
        }
    }

    /**
     * 填充对象
     */
    public static class ValueWrapper {
        /**
         * -1: EOF
         * 0: Null
         * 1: Object
         * 2: Map
         * 3: List map
         * 4: List Object
         */
        public int option;
        /**
         * 当{@code option}为3/4的时候，{@code i}表示list的消费下标
         */
        public int i;
        /**
         * 当{@code option}为3/4的时候，{@code size}表示已拉取数据大小
         */
        public int size;
        /**
         * 当{@code option=1}时，填充的数据保存到{@code o}中
         */
        public Object o;
        /**
         * 当{@code option=2}时，填充的数据保存到{@code map}中
         */
        public Map<String, Object> map;
        /**
         * 数据生产者，适用于数据量较大或长度未知的场景，效果等同于{@link ListSheet#more}方法
         */
        public BiFunction<Integer, Object, List<?>> supplier;
        /**
         * 当{@code option}为3/4的时候，填充的数据保存到{@code list}中
         */
        public List<Object> list;
        /**
         * 当{@code option}为1/4时，缓存对象的Field和Method方便后续取值
         */
        public Map<String, AccessibleObject> accessibleObjectMap;
    }
}

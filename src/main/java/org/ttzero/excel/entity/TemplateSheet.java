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

import org.slf4j.Logger;
import org.ttzero.excel.entity.e7.XMLWorksheetWriter;
import org.ttzero.excel.entity.style.Border;
import org.ttzero.excel.entity.style.Fill;
import org.ttzero.excel.entity.style.Font;
import org.ttzero.excel.entity.style.NumFmt;
import org.ttzero.excel.entity.style.Styles;
import org.ttzero.excel.manager.Const;
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
import java.beans.Introspector;
import java.beans.PropertyDescriptor;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.AccessibleObject;
import java.lang.reflect.Array;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Comparator;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.function.Supplier;

import static org.ttzero.excel.util.ReflectUtil.listDeclaredFields;

/**
 * 模板工作表，它支持指定一个已有的Excel文件作为模板导出，{@code TemplateSheet}将复制
 * 模板工作表的样式并替换占位符，同时{@code TemplateSheet}也可以和其它{@code Sheet}共用，
 * 这意味着可以添加多个模板工作表和普通工作表。需要注意的是多个模板可能产生重复的工作表名称，所以需要外部指定不同的名称以免
 * 打开文件异常
 *
 * <p>创建模板工作表需要指定模板文件，它可以是本地文件也可是输入流{@code InputStream}，支持的类型包含{@code xls}
 * 和{@code xlsx}两种格式，除模板文件外还需要指定Excel中的某个{@code Worksheet}工作表，
 * 未指定工作表时默认以第一个工作表做为模板，{@code TemplateSheet}工作表导出时不受{@code ExcelColumn}注解限制，
 * 导出的数据范围由模板内占位符决定</p>
 *
 * <p>默认占位符为一对关闭的大括号{@code ${key}}，可以使用{@link #setPrefix}和{@link #setSuffix}来重新指定占位符的前缀和后缀字符，
 * 建议不要设置太长的前后缀。占位符可以有一个命名空间，使用{@code ${namespace.key}}这种格式来添加命名空间</p>
 *
 * <p>使用{@link #setData}方法为占位符绑定值，如果指定了命名空间则绑定值时必须指定对应的命名空间，默认的命名空间为{@code null}，
 * 如果指定命名空间为 {@code this} 将被视为默认命名空间{@code null}，如果数据量较大时可绑定一个数据生产者{@link Supplier}
 * 来分片获取数据，它的作用与{@link ListSheet#more}方法一致。</p>
 *
 * <blockquote><pre>
 * new Workbook("模板测试")
 *      // 模板工作表
 *     .addSheet(new TemplateSheet(Paths.get("./template.xlsx")).setData(data))
 *     // 普通对象数组工作表
 *     .addSheet(new ListSheet&lt;&gt;())
 *     .writeTo("/tmp/");</pre></blockquote>
 *
 * @author guanquan.wang at 2023-12-01 15:10
 */
public class TemplateSheet extends Sheet {
    /**
     * 占位符前缀
     */
    protected String prefix = "${";
    /**
     * 占位符后缀
     */
    protected String suffix = "}";
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
     */
    protected int pf, pi;
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
     * @param o 任意对象，可以为Java Bean，Map，或者数组
     * @return 当前工作表
     */
    public TemplateSheet setData(String namespace, Object o) {
        if ("this".equals(namespace)) namespace = null;
        ValueWrapper vw = namespaceMapper.get(namespace);
        if (vw == null) {
            vw = new ValueWrapper();
            namespaceMapper.put(namespace, vw);
        }
        else LOGGER.warn("The namespace[{}] already exists.", namespace);
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
                    if (oo != null && vw.option == 0) vw.option = Map.class.isAssignableFrom(oo.getClass()) ? 3 : 4;
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
     * @param supplier 数据产生者
     * @return 当前工作表
     */
    public TemplateSheet setData(Supplier<List<?>> supplier) {
        return setData(null, supplier);
    }

    /**
     * 绑定一个{@code Supplier}到指定命名空间，适用于未知长度或数量最大的数组
     *
     * @param namespace 命名空间
     * @param supplier 数据产生者
     * @return 当前工作表
     */
    public TemplateSheet setData(String namespace, Supplier<List<?>> supplier) {
        if ("this".equals(namespace)) namespace = null;
        ValueWrapper vw = namespaceMapper.get(namespace);
        if (vw != null) {
            LOGGER.warn("The namespace[{}] already exists.", namespace);
        } else {
            vw = new ValueWrapper();
            namespaceMapper.put(namespace, vw);;
        }
        vw.supplier = supplier;

        // 加载第一批数据预处理数据类型
        if (supplier != null) {
            List list = supplier.get();
            Object oo = getFirstObject(list);
            if (oo != null) {
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
        // 清除数据（仅重置下标）
        rowBlock.clear();

        // 装载数据（这里不需要判断是否有表头，模板不需要表头）
        resetBlockData();

        // 使其可读
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
            for (; index < sheets.length && !originalSheetName.equals(sheets[index++].getName()); ) ;
            if (index > sheets.length)
                throw new IOException("The original sheet [" + originalSheetName + "] does not exist in template file.");
            originalSheetIndex = index - 1;
        }
        else if (originalSheetIndex < 0 || originalSheetIndex >= sheets.length)
            throw new IOException("The original sheet index [" + originalSheetIndex + "] is out of range in template file[0-" + sheets.length + "].");

        // 加载模板工作表
        FullSheet sheet = reader.sheet(originalSheetIndex).asFullSheet();

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

        // xlsx格式输出才进行以下格式复制
        if (writeAsExcel = sheetWriter != null && XMLWorksheetWriter.class.isAssignableFrom(sheetWriter.getClass())) {
            // 冻结,直接复制不需要计算移动
            Panes panes = sheet.getFreezePanes();
            if (panes != null) putExtProp(Const.ExtendPropertyKey.FREEZE, panes);

            // 合并
            List<Dimension> mergeCells0 = sheet.getMergeCells();
            if (mergeCells0 != null) {
                mergeCells = new ArrayList<>(mergeCells0.size());
                this.mergeCells0 = new HashMap<>(mergeCells0.size());
                // 这里将坐标切换到 base 0，方便后续取值
                for (Dimension dim : mergeCells0) this.mergeCells0.put(dimensionKey(dim), dim);
            }

            // 过滤
            Dimension autoFilter = sheet.getFilter();
            if (autoFilter != null) putExtProp(Const.ExtendPropertyKey.AUTO_FILTER, autoFilter);

            // 是否显示网格线
            this.showGridLines = sheet.isShowGridLines();

            // 预置列宽
            double defaultColWidth = sheet.getDefaultColWidth(), defaultRowHeight = sheet.getDefaultRowHeight();
            if (defaultColWidth >= 0) putExtProp("defaultColWidth", defaultColWidth);
            if (defaultRowHeight >= 0) putExtProp("defaultRowHeight", defaultRowHeight);

            // FIXME 图片（较为复杂不能简单复制，需要计算中间插入或扣除的行）
            List<Drawings.Picture> pictures = sheet.listPictures();
            if (pictures != null && !pictures.isEmpty()) {
                this.pictures = pictures.size() > 1 || !pictures.get(0).isBackground() ? new ArrayList<>(pictures) : null;
                for (Drawings.Picture p : pictures) {
                    if (p.isBackground()) setWaterMark(WaterMark.of(p.getLocalPath()));
                    else this.pictures.add(p);
                }
            }
        }

        // 预处理样式和占位符
        prepare();

        // 忽略表头输出
        super.ignoreHeader();
        // 初始化行迭代器
        rowIterator = new CommitRowSetIterator((RowSetIterator) sheet.reset().iterator());
        pf = preCells == null ? -1 : preCells[0][0].row;

        return len;
    }

    @Override
    protected void resetBlockData() {
        Integer xf;
        int len, n = 0, limit = sheetWriter.getRowLimit(); // 这里直接从writer中获取
        Dimension mergeCell;
        PreCell pn;
        Object e;
        Set<String> consumerValueKeys = null;
        for (int rbs = rowBlock.capacity(); n++ < rbs && rows < limit && rowIterator.hasNext(); ) {
            Row row = rowBlock.next();
            org.ttzero.excel.reader.Row row0 = rowIterator.next();
            // 设置行号
            row.index = rows = rowIterator.rows - 1;
            // 设置行高
            row.height = row0.getHeight();
            // 设置行是否隐藏
            row.hidden = row0.isHidden();
            // 空行特殊处理（lc-fc=-1)
            len = Math.max(row0.getLastColumnIndex() - row0.getFirstColumnIndex(), 0);
            Cell[] cells = row.realloc(len);
            // 预处理
            if (row0.getRowNum() == pf || rowIterator.hasFillCell) {
                if (!rowIterator.hasFillCell) rowIterator.withPreNodes(preCells[pi]);
                if (consumerValueKeys == null ) consumerValueKeys = new HashSet<>();
                else consumerValueKeys.clear();
            } else consumerValueKeys = null;

            // 占位符是否已消费结束
            boolean consumerEnd = true;

            for (int i = 0; i < len; i++) {
                Cell cell = cells[i], cell0 = row0.getCell(i);
                // Clear cells
                cell.clear();

                boolean fillCell = false;
                // 复制数据
                switch (row0.getCellType(cell0)) {
                    case STRING:
                        if (rowIterator.hasFillCell && (pn = rowIterator.preNodes[i]) != null) {
                            fillCell = true;
                            if (pn.nodes.length == 1) {
                                e = getNodeValue(pn.nodes[0]);
                                if (e != null) cellValueAndStyle.setCellValue(row, cell, e, new Column(), e.getClass(), false);
                                else cell.emptyTag();
                                consumerValueKeys.add(pn.nodes[0].namespace);
                            }
                            else {
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
                                    consumerValueKeys.add(node.namespace);
                                }
                                cell.setString(new String(pn.cb, 0, k));
                            }

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

                // 复制样式
                cell.xf = (xf = styleMap.get(cell0.xf)) != null ? xf : 0;

                // 合并单元格重新计算位置
                if (!fillCell && mergeCells0 != null && (mergeCell = mergeCells0.get(dimensionKey(row0.getRowNum() - 1, i))) != null) {
                    if (rows <= row0.getRowNum()) mergeCells.add(mergeCell);
                    else {
                        int r = rows - row0.getRowNum() + 1;
                        mergeCells.add(new Dimension(mergeCell.firstRow + r, mergeCell.firstColumn, mergeCell.lastRow + r, mergeCell.lastColumn));
                    }
                }
            }

            if (consumerValueKeys != null) {
                for (String vwKey : consumerValueKeys) {
                    ValueWrapper vw = namespaceMapper.get(vwKey);
                    // 如果为数组时需要移动游标
                    if (vw != null && (vw.option == 3 || vw.option == 4)) {
                        if (++vw.i < vw.list.size()) consumerEnd = false;
                            // 加载更多数据
                        else if (vw.supplier != null) {
                            List list = vw.supplier.get();
                            if (list != null && !list.isEmpty()) {
                                vw.list = list;
                                vw.i = 0;
                                consumerEnd = false;
                            } else vw.option = -1; // EOF
                        } else vw.option = -1; // EOF
                    }
                }
            }

            // 循环替换占位符时不要ark
            if (consumerEnd) {
                if (rowIterator.hasFillCell) {
                    pi++;
                    pf = preCells.length > pi && preCells[pi] != null &&  preCells[pi].length >= 1 ? preCells[pi][0].row : -1;
                }
                rowIterator.commit();
            }
        }
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
     * @param ao Method 或 Field
     * @param o 对象
     * @param logger 日志
     * @param key 占位符
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
    }

    @Override
    public void close() throws IOException {
        super.close();
        // 释放模板流
        if (reader != null) reader.close();
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
     */
    protected void prepare() {
        // 模板文件样式
        Styles styles0 = reader.getStyles(), styles = workbook.getStyles();
        // 样式缓存
        styleMap = new HashMap<>();
        int prefixLen = prefix.length(), suffixLen = suffix.length();
        for (Iterator<org.ttzero.excel.reader.Row> iter = reader.sheet(originalSheetIndex).iterator(); iter.hasNext(); ) {
            org.ttzero.excel.reader.Row row = iter.next();
            int index = 0;
            for (int i = row.getFirstColumnIndex(), end = row.getLastColumnIndex(); i < end; i++) {
                Cell cell = row.getCell(i);

                // 复制样式
                if (!styleMap.containsKey(cell.xf)) {
                    int style = row.getCellStyle(cell), xf = 0;
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
                    // 水平对齐
                    xf |= styles0.getHorizontal(style);
                    // 垂直对齐
                    xf |= styles0.getVertical(style);
                    // 自动折行
                    xf |= styles0.getWrapText(style);

                    // 添加进样式表
                    styleMap.put(cell.xf, styles.of(xf));
                }

                // 判断字符串是否包含占位符，可以是一个或多个
                if (row.getCellType(cell) == CellType.STRING) {
                    String v = row.getString(cell);
                    // 预处理单元格的值
                    PreCell preCell = prepareCellValue(v, prefixLen, suffixLen);
                    if (preCell != null) {
                        preCell.row = row.getRowNum();
                        preCell.col = i;

                        if (preCells == null) preCells = new PreCell[10][];
                        PreCell[] pns;
                        if (index == 0) {
                            if (pf >= preCells.length) preCells = Arrays.copyOf(preCells, preCells.length + 10);
                            preCells[pf++] = pns = new PreCell[Math.min(end - i, 10)];
                        } else if (index >= (pns = preCells[pf - 1]).length)
                            preCells[pf - 1] = pns = Arrays.copyOf(pns, pns.length + 10);
                        pns[index++] = preCell;

                        // 检测是否包含单行合并单元格
                        if (mergeCells0 != null) {
                            Dimension mc = mergeCells0.get(dimensionKey(row.getRowNum() - 1, i));
                            // FIXME 目前只支持单行合并
                            if (mc != null && mc.height == 1) preCell.m = mc.width - 1;
                        }
                    }
                }
            }

            if (index > 0 && preCells[pf - 1].length > index) preCells[pf - 1] = Arrays.copyOf(preCells[pf - 1], index);
        }
    }

    /**
     * 单元格字符串预处理，检测是否包含占位符以及占位符预处理
     *
     * @param v 单元格原始值
     * @param prefixLen 占位符前缀长度
     * @param suffixLen 占位符后缀长度
     * @return 预处理结果
     */
    protected PreCell prepareCellValue(String v, int prefixLen, int suffixLen) {
        int len = v.length();
        if (len <= prefixLen + suffixLen) return null;
        int pi = 0, fi = v.indexOf(prefix);
        if (fi < 0) return null;
        int fn, li = v.indexOf(suffix, fn = fi + prefixLen);
        if (li <= fn) return null;

        PreCell pn = new PreCell();
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
                node.namespace = v.substring(fn, j);
                node.val = v.substring(j + 1, li);
            } else node.val = v.substring(fn, li);

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
        return pn;
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
            PropertyDescriptor[] propertyDescriptors = Introspector.getBeanInfo(clazz, Object.class)
                .getPropertyDescriptors();
            for (PropertyDescriptor pd : propertyDescriptors) {
                Method method = pd.getReadMethod();
                if (method != null) tmp.put(pd.getName(), method);
            }

            Field[] declaredFields = listDeclaredFields(clazz);
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
     * 将范围首行首列进行计算后转为缓存的Key
     *
     * @param dim 范围{@link Dimension}
     * @return 缓存Key
     */
    public static long dimensionKey(Dimension dim) {
        return ((dim.firstColumn - 1) & 0x7FFF) | ((long) dim.firstRow - 1) << 16;
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
     * 需要手动调用{@link #commit()}方法才会移动游标
     */
    public static class CommitRowSetIterator implements Iterator<org.ttzero.excel.reader.Row> {
        public final RowSetIterator iterator;
        public org.ttzero.excel.reader.Row current;
        public int rows;
        public PreCell[] preNodes;
        public boolean hasFillCell;

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

        public void withPreNodes(PreCell[] preNodes) {
            int len = current.getLastColumnIndex(), len0 = preNodes[preNodes.length - 1].col + 1;
            this.preNodes = new PreCell[Math.max(len, len0)];
            for (PreCell p : preNodes) this.preNodes[p.col] = p;
            hasFillCell = true;
        }

        /**
         * 提交后才将移动到下一行，否则一直停留在当前行
         */
        public void commit() {
            current = null;
            preNodes = null;
            hasFillCell = false;
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
         * 合并范围 正数为行合并 负数为列合并
         */
        public Integer m;
    }

    /**
     * 单元格预处理节点
     */
    public static class Node {
        /**
         * 标志位集合，保存一些简单的标志位以节省空间，对应的位点说明如下
         *
         * <blockquote><pre>
         *  Bit  | Contents
         * ------+---------
         * 0, 1 | 是否为占位符 0: 普通文本 1: 占位符
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

        @Override
        public String toString() {
            return (option & 1) == 0 ? val : ("$" + (namespace != null ? namespace + "::" : "") + val);
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
        public Supplier<List<?>> supplier;
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

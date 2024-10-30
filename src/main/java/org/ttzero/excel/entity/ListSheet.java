/*
 * Copyright (c) 2017-2018, guanquan.wang@yandex.com All Rights Reserved.
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

import org.ttzero.excel.annotation.ExcelColumn;
import org.ttzero.excel.annotation.ExcelColumns;
import org.ttzero.excel.annotation.FreezePanes;
import org.ttzero.excel.annotation.HeaderComment;
import org.ttzero.excel.annotation.HeaderStyle;
import org.ttzero.excel.annotation.Hyperlink;
import org.ttzero.excel.annotation.IgnoreExport;
import org.ttzero.excel.annotation.MediaColumn;
import org.ttzero.excel.annotation.StyleDesign;
import org.ttzero.excel.drawing.PresetPictureEffect;
import org.ttzero.excel.manager.Const;
import org.ttzero.excel.processor.ConversionProcessor;
import org.ttzero.excel.processor.Converter;
import org.ttzero.excel.processor.StyleProcessor;
import org.ttzero.excel.reader.Cell;
import org.ttzero.excel.util.StringUtil;

import java.beans.IntrospectionException;
import java.io.IOException;
import java.lang.reflect.AccessibleObject;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.lang.reflect.ParameterizedType;
import java.lang.reflect.Type;
import java.util.ArrayList;
import java.util.Collection;
import java.util.Collections;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.function.BiFunction;

import static org.ttzero.excel.util.ReflectUtil.listDeclaredFieldsUntilJavaPackage;
import static org.ttzero.excel.util.ReflectUtil.listReadMethods;
import static org.ttzero.excel.util.ReflectUtil.readMethodsMap;
import static org.ttzero.excel.util.StringUtil.EMPTY;
import static org.ttzero.excel.util.StringUtil.isEmpty;
import static org.ttzero.excel.util.StringUtil.isNotEmpty;

/**
 * 对象数组工作表，内部使用{@code List<T>}做数据源，所以它是应用最广泛的一种工作表。
 * {@code ListSheet}默认支持数据切片，少量数据可以在实例化时一次性传入，数据量较大时建议切片获取数据
 *
 * <pre>
 * new Workbook("11月待拜访客户")
 *     .addSheet(new ListSheet&lt;Customer&gt;() {
 *         &#x40;Override
 *         protected List&lt;Customer&gt; more() {
 *             // 分页查询数据
 *             List&lt;Customer&gt; list = customerService.list(queryVo);
 *             // 页码 + 1
 *             queryVo.setPageNum(queryVo.getPageNum() + 1);
 *             return list;
 *         }
 *     }).writeTo(response.getOutputStream());</pre>
 *
 * <p>如上示例覆写{@link #more}方法获取切片数据，直到返回空数据或{@code null}为止,这样不至少将大量数据堆积到内存，
 * 输出协议使用{@link RowBlock}进行装填数据并落盘。{@code more}方法在{@code ListSheet}工作表是一定会被
 * 调用的即使初始化工作表传入了数据，工作表判断无需导出的情况除外，比如未指定表头且{@code Bean}对象无任何&#x40;ExcelColumn注释，
 * 则会导出空工作表</p>
 *
 * <p>{@code ListSheet}使用{@link #getTClass}方法获取泛型的实际类型，内部优先使用{@link Class#getGenericSuperclass}方法获取，
 * 如果有子类指定{@code T}类型则可以获取到{@code T}的类型，否则将使用数组中第一条数据做为泛型的具体类型，
 * 如果希望无数据时依然导出表头就必须让{@code ListSheet}获取泛型{@code T}的实际类型，当无数据且无子类指定{@code T}时可以使用
 * {@link #setClass}方法设置泛型的类型</p>
 *
 * <p>大多数使用{@code ListSheet}工作表导数据时都会使用注解，{@code ListSheet}支持自定义注解以延申功能，
 * 你甚至可以不使用任何的内置注释全部使用自定义注解来实现导入导出，使用自定义注解时需要搭配自定义{@code ListSheet}解析注解，
 * 其中最重要的两个方法{@link #ignoreColumn(AccessibleObject)}和{@link #createColumn(AccessibleObject)}就是读取
 * 方法或字段上的注解来创建{@code Column}对象，{@code AccessibleObject}可能是一个{@code Method}或者一个{@code Field}</p>
 *
 * <p>对象取值会优先调用get方法获取，如果未发现get方法则直接从{@code field}取值。导出数据并不仅限于get方法，
 * 你可以在任何无参且仅有一个返回值的方法上添加&#x40;ExcelColumn注解进行导出，你还可以在子类定义相同的方法来替换父类上的&#x40;ExcelColumn注解，
 * 解析注解时会从子类往上逐级解析至到父级为{@code Object}为止</p>
 *
 * <p>除子类覆写{@link #more}方法外还可以通过{@link #setData(BiFunction)}设置一个数据生产者，它可以减化数据分片的代码。
 * {@code dataSupplier}被定义为{@code BiFunction<Integer, T, List<T>>}，其中第一个入参{@code Integer}表示已拉取数据的记录数
 * （并非已写入数据），第二个入参{@code T}表示上一批数据中最后一个对象，业务端可以通过这两个参数来计算下一批数据应该从哪个节点开始拉取，
 * 通常你可以使用第一个参数除以每批拉取的数据大小来确定当前页码，如果数据已排序则可以使用{@code T}对象的排序字段来计算下一批数据的游标从而跳过
 * {@code limit ... offset ... }分页查询从而极大提升取数性能</p>
 *
 * <pre>
 * new Workbook()
 *     .addSheet(new ListSheet&lt;Customer&gt;()
 *         // 分页查询，每页查询100条数据，可以通过已拉取记录数计算当前页面
 *         .setData((i, lastOne) -&gt; customerService.pagingQuery(i/100, 100))
 *     ).writeTo(Paths.get("f://abc.xlsx"));</pre>
 *
 * <p>参考文档:</p>
 * <p><a href="https://github.com/wangguanquan/eec/wiki/%E9%AB%98%E7%BA%A7%E7%89%B9%E6%80%A7#%E8%87%AA%E5%AE%9A%E4%B9%89%E6%B3%A8%E8%A7%A3">自定义注解</a></p>
 *
 * @author guanquan.wang at 2018-01-26 14:48
 * @see ListMapSheet
 * @see SimpleSheet
 */
public class ListSheet<T> extends Sheet {
    /**
     * 临时存放数据
     */
    protected List<T> data;
    /**
     * 控制读取开始和结束下标
     */
    protected int start, end;
    /**
     * 结束标记{@code EOF}
     */
    protected boolean eof;
    /**
     * 泛型&lt;T&gt;的实际类型
     */
    protected Class<?> tClazz;
    /**
     * 行级动态样式处理器
     */
    protected StyleProcessor<T> styleProcessor;
    /**
     * 强制导出，忽略安全限制全字段导出，确认需求后谨慎使用
     */
    protected int forceExport;
    /**
     * 数据产生者，简化分片查询
     */
    protected BiFunction<Integer, T, List<T>> dataSupplier;

    /**
     * 设置行级动态样式处理器，作用于整行优先级高于单元格动态样式处理器
     *
     * @param styleProcessor 行级动态样式处理器
     * @return 当前工作表
     */
    public Sheet setStyleProcessor(StyleProcessor<T> styleProcessor) {
        this.styleProcessor = styleProcessor;
        putExtProp(Const.ExtendPropertyKey.STYLE_DESIGN, styleProcessor);
        return this;
    }

    /**
     * 获取当前工作表的行级动态样式处理器，如果未设置则从扩展参数中查找
     *
     * @return 行级动态样式处理器
     */
    public StyleProcessor<T> getStyleProcessor() {
        if (styleProcessor != null) return styleProcessor;
        @SuppressWarnings("unchecked")
        StyleProcessor<T> fromExtProp = (StyleProcessor<T>) getExtPropValue(Const.ExtendPropertyKey.STYLE_DESIGN);
        return this.styleProcessor = fromExtProp;
    }

    /**
     * 实例化工作表，未指定工作表名称时默认以{@code 'Sheet'+id}命名
     */
    public ListSheet() {
        super();
    }

    /**
     * 实例化工作表并指定工作表名称
     *
     * @param name 工作表名称
     */
    public ListSheet(String name) {
        super(name);
    }

    /**
     * 实例化工作表并指定表头信息
     *
     * @param columns 表头信息
     */
    public ListSheet(final Column... columns) {
        super(columns);
    }

    /**
     * 实例化工作表并指定工作表名称和表头信息
     *
     * @param name    工作表名称
     * @param columns 表头信息
     */
    public ListSheet(String name, final Column... columns) {
        super(name, columns);
    }

    /**
     * 实例化工作表并指定工作表名称，水印和表头信息
     *
     * @param name      工作表名称
     * @param waterMark 水印
     * @param columns   表头信息
     */
    public ListSheet(String name, WaterMark waterMark, final Column... columns) {
        super(name, waterMark, columns);
    }

    /**
     * 实例化工作表并指定初始数据
     *
     * @param data 初始数据
     */
    public ListSheet(List<T> data) {
        this(null, data);
    }

    /**
     * 实例化工作表并指定工作表名称和初始数据
     *
     * @param name 工作表名称
     * @param data 初始数据
     */
    public ListSheet(String name, List<T> data) {
        super(name);
        setData(data);
    }

    /**
     * 实例化工作表并指定初始数据和表头
     *
     * @param data    初始数据
     * @param columns 表头信息
     */
    public ListSheet(List<T> data, final Column... columns) {
        this(null, data, columns);
    }

    /**
     * 实例化工作表并指定工作表名称、初始数据和表头
     *
     * @param name    工作表名称
     * @param data    初始数据
     * @param columns 表头信息
     */
    public ListSheet(String name, List<T> data, final Column... columns) {
        this(name, data, null, columns);
    }

    /**
     * 实例化工作表并指定初始数据、水印和表头
     *
     * @param data      初始数据
     * @param waterMark 水印
     * @param columns   表头信息
     */
    public ListSheet(List<T> data, WaterMark waterMark, final Column... columns) {
        this(null, data, waterMark, columns);
    }

    /**
     * 实例化工作表并指定工作表名称、初始数据、水印和表头
     *
     * @param name      工作表名称
     * @param data      初始数据
     * @param waterMark 水印
     * @param columns   表头信息
     */
    public ListSheet(String name, List<T> data, WaterMark waterMark, final Column... columns) {
        super(name, waterMark, columns);
        setData(data);
    }

    /**
     * 指定泛型{@code T}的实际类型，不指定时默认由反射或数组中第一个对象类型而定
     *
     * @param tClazz 泛型{@code T}的实际类型
     * @return 当前工作表
     */
    public Sheet setClass(Class<?> tClazz) {
        this.tClazz = tClazz;
        return this;
    }

    /**
     * 设置初始数据，导出的时候依然会调用{@link #more()} 方法以获取更多数据
     *
     * @param data 初始数据
     * @return 当前工作表
     */
    public ListSheet<T> setData(final List<T> data) {
        if (data == null) return this;
        this.data = new ArrayList<>(data);
        // Has data and worksheet can write
        // Paging in advance
        if (sheetWriter != null) {
            paging();
        }
        return this;
    }

    /**
     * 设置数据生产者，如果设置了此值{@link #more}方法将从此生产者中获取数据
     *
     * @param dataSupplier 数据生产者其中{@code Integer}为已拉取数据的记录数，{@code T}为上一批数据中最后一个对象
     * @return 当前工作表
     */
    public ListSheet<T> setData(BiFunction<Integer, T, List<T>> dataSupplier) {
        this.dataSupplier = dataSupplier;
        return this;
    }

    /**
     * 获取队列中第一个非{@code null}对象用于解析
     *
     * @return 第一个非 {@code null}对象
     */
    protected T getFirst() {
        // 初始没有数据时调用一次more方法获取数据
        if (data == null || data.isEmpty()) {
            List<T> more = more();
            if (more != null && !more.isEmpty()) data = new ArrayList<>(more);
            else return null;
        }
        T first = data.get(start);
        if (first != null) return first;
        int i = start + 1;
        do {
            first = data.get(i++);
        } while (first == null && i< this.data.size());
        return first;
    }

    /**
     * Release resources
     *
     * @throws IOException if I/O error occur
     */
    @Override
    public void close() throws IOException {
        // Maybe there has more data
        if (!eof && rows >= getRowLimit()) {
            List<T> list = more();
            if (list != null && !list.isEmpty()) {
                compact();
                data.addAll(list);
                @SuppressWarnings("unchecked")
                ListSheet<T> copy = getClass().cast(clone());
                copy.start = 0;
                copy.end = list.size();
                workbook.insertSheet(id, copy);
                // Do not close current worksheet
                shouldClose = false;
            }
        }
        if (shouldClose && data != null) {
            // Some Collection not support #remove
//            data.clear();
            data = null;
        }
        super.close();
    }

    /**
     * 重置{@code RowBlock}行块数据
     */
    @Override
    protected void resetBlockData() {
        if (!eof && left() < rowBlock.capacity()) {
            append();
        }

        // Find the end index of row-block
        int end = getEndIndex(), len = columns.length;
        boolean hasGlobalStyleProcessor = (extPropMark & 2) == 2;
        try {
            for (; start < end; rows++, start++) {
                Row row = rowBlock.next();
                row.index = rows;
                Cell[] cells = row.realloc(len);
                T o = data.get(start);
                boolean isNull = o == null;
                for (int i = 0; i < len; i++) {
                    // Clear cells
                    Cell cell = cells[i];
                    cell.clear();

                    Object e;
                    EntryColumn column = (EntryColumn) columns[i];
                    /*
                    The default processing of null values still retains the row style.
                    If you don't want any style and value, you can change it to {@code continue}
                     */
                    if (column.isIgnoreValue() || isNull)
                        e = null;
                    else {
                        if (column.getMethod() != null)
                            e = column.getMethod().invoke(o);
                        else if (column.getField() != null)
                            e = column.getField().get(o);
                        else e = o;
                    }

                    cellValueAndStyle.reset(row, cell, e, column);
                    if (hasGlobalStyleProcessor) {
                        cellValueAndStyle.setStyleDesign(o, cell, column, getStyleProcessor());
                    }
                }
                row.height = getRowHeight();
            }
        } catch (IllegalAccessException | InvocationTargetException e) {
            throw new ExcelWriteException(e);
        }
    }

    /**
     * 加载数据，内部调用{@link #more}获取数据并判断是否需要分页，超过工作表行上限则调用{@link #paging}分页
     */
    protected void append() {
        int rbs = rowBlock.capacity();
        for (; ; ) {
            List<T> list = more();
            // No more data
            if (list == null || list.isEmpty()) {
                eof = shouldClose = true;
                break;
            }
            // The first getting
            if (data == null) {
                setData(list);

                if (list.size() < rbs) continue;
                else break;
            }
            compact();
            data.addAll(list);
            start = 0;
            end = data.size();
            // Split worksheet
            if (end >= rbs) {
                paging();
                break;
            }
        }
    }

    private void compact() {
        // Copy the remaining data to a temporary array
        int size = left();
        if (start > 0 && size > 0) {
            // append and resize
            List<T> last = new ArrayList<>(size);
            last.addAll(data.subList(start, end));
            data.clear();
            data.addAll(last);
        } else if (start > 0) data.clear();
    }

    /**
     * 获取泛型T的实际类型，优先使用{@link Class#getGenericSuperclass}方法获取，如果有子类指定{@code T}类型则可以获取，
     * 否则将使用数组中第一条数据做为泛型的具体类型
     *
     * @return T的实际类型，无法获取具体类型时返回{@code null}
     */
    protected Class<?> getTClass() {
        Class<?> clazz = tClazz;
        if (clazz != null) return clazz;
        if (getClass().getGenericSuperclass() instanceof ParameterizedType) {
            Type type = ((ParameterizedType) getClass()
                .getGenericSuperclass()).getActualTypeArguments()[0];
            if (type instanceof Class) {
                clazz = (Class) type;
            }
        }
        if (clazz == null) {
            T o = getFirst();
            if (o == null) return null;
            clazz = o.getClass();
        }
        tClazz = clazz;
        return clazz;
    }

    /**
     * 初始化表头信息，如果未指定{@code Columns}则默认反射{@code T}及其父类的所有字段，
     * 并采集所有标记有{@link ExcelColumn}注解的字段和方法（这里限制无参数且仅有一个返回值的方法），
     * {@code Column}顺序由{@code colIndex}决定，如果没有{@code colIndex}则按字段和方法在
     * Java Bean中的定义顺序而定。
     *
     * <p>如果有指定{@code Columns}则忽略排序仅将{@link Column#key}与字段和方法进行绑定方便后续取值</p>
     *
     * @return 表头列个数
     */
    protected int init() {
        Class<?> clazz = getTClass();
        if (clazz == null) return columns != null ? columns.length : 0;

        Map<String, Method> tmp = new HashMap<>();
        try {
            tmp.putAll(readMethodsMap(clazz, Object.class));
        } catch (IntrospectionException e) {
            LOGGER.warn("Get class {} methods failed.", clazz);
        }

        Field[] declaredFields = listDeclaredFieldsUntilJavaPackage(clazz, c -> !ignoreColumn(c));

        boolean forceExport = this.forceExport == 1;

        if (!hasHeaderColumns()) {
            // Get ExcelColumn annotation method
            List<Column> list = new ArrayList<>(declaredFields.length);
            Map<String, Method> existsMethod = new HashMap<>(declaredFields.length);
            for (int i = 0; i < declaredFields.length; i++) {
                Field field = declaredFields[i];
                field.setAccessible(true);
                String gs = field.getName();

                // Ignore annotation on read method
                Method method = tmp.get(gs);
                if (method != null) {
                    existsMethod.put(gs, method);
                    // Filter all ignore column
                    if (ignoreColumn(method)) {
                        declaredFields[i] = null;
                        continue;
                    }

                    EntryColumn column = createColumn(method);
                    // Force export
                    if (column == null && forceExport) {
                        column = new EntryColumn(gs, EMPTY, false);
                    }
                    if (column != null) {
                        EntryColumn tail = (EntryColumn) column.getTail();
                        tail.method = method;
                        tail.field = field;
                        tail.clazz = method.getReturnType();
                        tail.key = gs;
                        if (isEmpty(tail.name)) {
                            tail.name = gs;
                        }
                        list.add(column);

                        // Attach header style
                        buildHeaderStyle(method, field, tail);
                        // Attach header comment
                        buildHeaderComment(method, field, tail);
                        continue;
                    }
                }

                EntryColumn column = createColumn(field);
                // Force export
                if (column == null && forceExport) {
                    column = new EntryColumn(gs, EMPTY, false);
                }
                if (column != null) {
                    list.add(column);
                    EntryColumn tail = (EntryColumn) column.getTail();
                    tail.field = field;
                    tail.key = gs;
                    if (isEmpty(tail.name)) {
                        tail.name = gs;
                    }
                    if (method != null) {
                        tail.clazz = method.getReturnType();
                        tail.method = method;
                    } else tail.clazz = field.getType();

                    // Attach header style
                    buildHeaderStyle(method, field, tail);
                    // Attach header comment
                    buildHeaderComment(method, field, tail);
                }
            }

            // Attach some custom column
            List<Column> attachList = attachOtherColumn(existsMethod, clazz);
            if (attachList != null) list.addAll(attachList);

            // No column to write
            if (list.isEmpty()) {
                // 如果没有列可写则判断是否为简单类型，简单类型可直接输出
                if (cellValueAndStyle.isAllowDirectOutput(clazz)) {
                    list.add(new EntryColumn().setClazz(clazz));
                } else {
                    headerReady = eof = shouldClose = true;
                    this.end = 0;
                    if (java.util.Map.class.isAssignableFrom(clazz))
                        LOGGER.warn("List<Map> has detected, please use ListMapSheet for export.");
                    else LOGGER.warn("Class [{}] do not contains properties to export.", clazz);
                    return 0;
                }
            }
            columns = new Column[list.size()];
            list.toArray(columns);
        } else {
            Method[] others = filterOthersMethodsCanExport(Collections.emptyMap(), clazz);
            Map<String, Method> otherMap = new HashMap<>();
            for (Method m : others) {
                ExcelColumn ec = m.getAnnotation(ExcelColumn.class);
                if (ec != null && StringUtil.isNotEmpty(ec.value())) {
                    otherMap.put(ec.value(), m);
                }
                otherMap.put(m.getName(), m);
            }
            for (int i = 0; i < columns.length; i++) {
                Column hc = new EntryColumn(columns[i]);
                columns[i] = hc;
                if (hc.tail != null) {
                    hc = hc.tail;
                }

                // If key miss
                if (StringUtil.isEmpty(hc.key)) hc.key = hc.name;

                EntryColumn ec = (EntryColumn) hc;
                if (ec.method == null) {
                    Method method = tmp.get(hc.key);
                    if (method != null) {
                        ec.method = method;
                    } else if ((method = otherMap.get(hc.key)) != null) {
                        ec.method = method;
                    }
                }

                if (ec.field == null) {
                    for (Field field : declaredFields) {
                        if (field.getName().equals(hc.key)) {
                            field.setAccessible(true);
                            ec.field = field;
                            break;
                        }
                    }
                }

                if (ec.method == null && ec.field == null) {
                    if (columns.length > 1) {
                        LOGGER.warn("Column [" + hc.getName() + "(" + hc.key + ")"
                            + "] not declare in class " + clazz);
                        hc.ignoreValue();
                    }
                    // Write as Object#toString()
                    else LOGGER.warn("Column one does not specify method or filed");
                } else if (hc.getClazz() == null) {
                    hc.setClazz(ec.method != null ? ec.method.getReturnType() : ec.field.getType());
                }

                // Attach header style
                if (hc.getHeaderStyleIndex() == -1) {
                    buildHeaderStyle(ec.method, ec.field, hc);
                }
                // Attach header comment
                if (hc.headerComment == null) {
                    buildHeaderComment(ec.method, ec.field, hc);
                }
            }
        }

        // Merge Header Style defined on Entry Class
        mergeGlobalSetting(clazz);

        return columns.length;
    }

    /**
     * 创建列信息，默认解析&#x40;Comment注解，支持自定义注解
     *
     * @param ao {@link AccessibleObject} 实体中的属性或方法
     * @return 单列表头信息
     */
    protected EntryColumn createColumn(AccessibleObject ao) {
        // Filter all ignore column
        if (ignoreColumn(ao)) return null;

        ao.setAccessible(true);
        // Style Design
        StyleProcessor<?> sp = getDesignStyle(ao);

        EntryColumn root = null;
        // Support multi header columns
        ExcelColumns cs = ao.getAnnotation(ExcelColumns.class);
        if (cs != null) {
            ExcelColumn[] ecs = cs.value();
            for (ExcelColumn ec : ecs) {
                EntryColumn column = createColumnByAnnotation(ec);
                if (sp != null) {
                    column.styleProcessor = sp;
                }
                if (root == null) {
                    root = column;
                } else {
                    root.addSubColumn(column);
                }
            }
        }
        // Single header column
        else {
            ExcelColumn ec = ao.getAnnotation(ExcelColumn.class);
            if (ec != null) {
                root = createColumnByAnnotation(ec);
                if (sp != null) {
                    root.styleProcessor = sp;
                }
            }
        }

        MediaColumn mediaColumn = ao.getAnnotation(MediaColumn.class);
        if (mediaColumn != null) {
            if (root == null) root = new EntryColumn(" ", EMPTY, false);
            Column tail = root.getTail();
            tail.writeAsMedia();
            if (mediaColumn.presetEffect() != PresetPictureEffect.None) {
                tail.setEffect(mediaColumn.presetEffect().getEffect());
            }
        }
        // Hyperlink
        else if (root != null) {
            Hyperlink Hyperlink = ao.getAnnotation(Hyperlink.class);
            if (Hyperlink != null) root.getTail().writeAsHyperlink();
        }
        return root;
    }

    /**
     * 解析&#x40;ExcelColumn注解并创建表头
     *
     * @param ec {@code ExcelColumn}注解
     * @return 单列表头信息
     */
    protected EntryColumn createColumnByAnnotation(ExcelColumn ec) {
        if (ec == null) return null;
        EntryColumn column = new EntryColumn(ec.value(), EMPTY, ec.share());
        // Number format
        if (isNotEmpty(ec.format())) {
            column.setNumFmt(ec.format());
        }
        // Wrap
        column.setWrapText(ec.wrapText());
        // Column index
        if (ec.colIndex() > -1) {
            column.colIndex = ec.colIndex();
        }
        // Hidden Column
        if (ec.hide()) column.hide();
        // Cell max width
        if (ec.maxWidth() >= 0.0D) column.width = ec.maxWidth();
        // Converter
        if (!Converter.None.class.isAssignableFrom(ec.converter())) {
            try {
                 column.setConverter(ec.converter().getDeclaredConstructor().newInstance());
            } catch (InstantiationException | IllegalAccessException | NoSuchMethodException | InvocationTargetException e) {
                LOGGER.warn("Construct {} error occur, it will be ignore.", ec.converter(), e);
            }
        }

        return column;
    }

    /**
     * 构建样式，默认解析{@link HeaderStyle}注解
     *
     * <p>优选从方法上获取注解，如果没有则从field中获取</p>
     *
     * @param main   Method
     * @param sub    Field
     * @param column 需要添加样式的表头
     */
    protected void buildHeaderStyle(AccessibleObject main, AccessibleObject sub, Column column) {
        HeaderStyle hs = null;
        if (main != null) {
            hs = main.getAnnotation(HeaderStyle.class);
        }
        if (hs == null && sub != null) {
            hs = sub.getAnnotation(HeaderStyle.class);
        }
        if (hs != null) {
            column.setHeaderStyle(this.buildHeadStyle(hs.fontColor(), hs.fillFgColor()));
        }
    }

    /**
     * 构建“批注”，默认解析{@link HeaderComment}注解，支持自定义注解
     *
     * <p>优选从方法上获取注解，如果没有则从field中获取</p>
     *
     * @param main   Method
     * @param sub    Field
     * @param column 需要添加批注的表头
     */
    protected void buildHeaderComment(AccessibleObject main, AccessibleObject sub, Column column) {
        HeaderComment comment = null;
        if (main != null) {
            comment = main.getAnnotation(HeaderComment.class);
            if (comment == null) {
                ExcelColumn ec = main.getAnnotation(ExcelColumn.class);
                if (ec != null) comment = ec.comment();
            }
        }
        if (comment == null && sub != null) {
            comment = sub.getAnnotation(HeaderComment.class);
            if (comment == null) {
                ExcelColumn ec = sub.getAnnotation(ExcelColumn.class);
                if (ec != null) comment = ec.comment();
            }
        }
        if (comment != null && (isNotEmpty(comment.value()) || isNotEmpty(comment.title()))) {
            column.headerComment = new Comment(comment.title(), comment.value(), comment.width(), comment.height());
        }
    }

    /**
     * 设置全局设置
     *
     * @param clazz Class of &lt;T&gt;
     */
    protected void mergeGlobalSetting(Class<?> clazz) {
        HeaderStyle headerStyle = clazz.getDeclaredAnnotation(HeaderStyle.class);
        int style = 0;
        if (headerStyle != null) {
            style = buildHeadStyle(headerStyle.fontColor(), headerStyle.fillFgColor());
        }
        for (Column column : columns) {
            // 如果字段未独立设置样式则使用方法上的样式
            if (style > 0 && column.getHeaderStyleIndex() == -1 && column.headerStyle == null)
                column.setHeaderStyle(style);
        }

        // Parse the global style processor
        if (styleProcessor == null) {
            @SuppressWarnings({"unchecked", "retype"})
            StyleProcessor<T> styleProcessor = (StyleProcessor<T>) getDesignStyle(clazz);
            if (styleProcessor != null) setStyleProcessor(styleProcessor);
        }

        // Freeze panes
        attachFreezePanes(clazz);
    }

    /**
     * 读取类上的样式处理器注解，默认解析{@link StyleDesign}注解，支持自定义注解
     *
     * @param clazz Class of &lt;T&gt;
     * @return 样式处理器或 {@code null}
     */
    protected StyleProcessor<?> getDesignStyle(Class<?> clazz) {
        StyleDesign styleDesign = clazz.getDeclaredAnnotation(StyleDesign.class);
        return getDesignStyle(styleDesign);
    }

    /**
     * 读取单个字段或者方法上的样式处理器注解，默认解析{@link StyleDesign}注解，支持自定义注解
     *
     * @param ao 字段{@code Field}或方法{@code Method}
     * @return 样式处理器或 {@code null}
     */
    protected StyleProcessor<?> getDesignStyle(AccessibleObject ao) {
        StyleDesign styleDesign = ao.getAnnotation(StyleDesign.class);
        return getDesignStyle(styleDesign);
    }

    /**
     * 读取样式处理器
     *
     * @param styleDesign 默认{@link StyleDesign}
     * @return 样式处理器
     */
    protected StyleProcessor<?> getDesignStyle(StyleDesign styleDesign) {
        if (styleDesign != null && !StyleProcessor.None.class.isAssignableFrom(styleDesign.using())) {
            try {
                return styleDesign.using().getDeclaredConstructor().newInstance();
            } catch (InstantiationException | IllegalAccessException | NoSuchMethodException | InvocationTargetException e) {
                LOGGER.warn("Construct {} error occur, it will be ignore.", styleDesign.using(), e);
            }
        }
        return null;
    }

    /**
     * 导出时忽略某些字段或方法，默认解析{@link IgnoreExport}注解，支持自定义注解
     *
     * @param ao {@code Method} or {@code Field}
     * @return 如果忽略该属性返回true
     */
    protected boolean ignoreColumn(AccessibleObject ao) {
        return ao.getAnnotation(IgnoreExport.class) != null;
    }

    /**
     * 收集可导出的{@code Method}并创建Column对象
     *
     * @param existsMethodMapper 已有的可导出{@code Method}映射，key为方法名
     * @param clazz              Class of &lt;T&gt;
     * @return 新收集的可导出{@code Column}数组
     */
    protected List<Column> attachOtherColumn(Map<String, Method> existsMethodMapper, Class<?> clazz) {
        // Collect the method which has ExcelColumn annotation
        Method[] readMethods = filterOthersMethodsCanExport(existsMethodMapper, clazz);

        if (readMethods != null) {
            Set<Method> existsMethods = new HashSet<>(existsMethodMapper.values());
            List<Column> list = new ArrayList<>();
            for (Method method : readMethods) {
                // Exclusions exists
                if (existsMethods.contains(method)) continue;
                EntryColumn column = createColumn(method);
                if (column != null) {
                    list.add(column);
                    EntryColumn tail = (EntryColumn) column.getTail();
                    tail.method = method;
                    tail.clazz = method.getReturnType();
                    tail.key = method.getName();
                    if (isEmpty(tail.name)) {
                        tail.name = method.getName();
                        if (tail.name.startsWith("get")) tail.name = StringUtil.lowFirstKey(tail.name.substring(3));
                        else if (tail.name.startsWith("is")) tail.name = StringUtil.lowFirstKey(tail.name.substring(2));
                    }

                    // Attach header style
                    buildHeaderStyle(method, null, tail);
                    // Attach header comment
                    buildHeaderComment(method, null, tail);
                }
            }
            return list;
        }
        return null; // No more columns
    }

    /**
     * 获取表头信息，未实例化表头时会执行初始化方法实例化表头
     *
     * @return 表头信息
     */
    @Override
    protected Column[] getHeaderColumns() {
        if (!headerReady) {
            // create header columns
            int size = init();
            if (size <= 0) {
                columns = new Column[0];
            }
        }
        return columns;
    }

    /**
     * 计算需要读取的尾下标，相对于当前数组的尾下标
     *
     * @return 尾下标
     */
    protected int getEndIndex() {
        int blockSize = rowBlock.capacity(), rowLimit = getRowLimit();
        if (rows + blockSize > rowLimit) {
            blockSize = rowLimit - rows;
        }
        int end = start + blockSize;
        return Math.min(end, this.end);
    }

    /**
     * 数组中剩余多少数据
     *
     * @return 数组中剩余数据量
     */
    protected int left() {
        return end - start;
    }

    /**
     * 分页处理，如果达到分页条件时会复制一个新的工作表插入到当前位置之后
     */
    @Override
    protected void paging() {
        int len = dataSize(), limit = getRowLimit();
        // paging
        if (len + rows > limit) {
            // Reset current index
            end = limit - rows + start;
            shouldClose = false;
            eof = true;

            int n = id;
            for (int i = end; i < len; ) {
                @SuppressWarnings("unchecked")
                ListSheet<T> copy = getClass().cast(clone());
                copy.start = i;
                copy.end = (i = Math.min(i + limit, len));
                copy.eof = copy.end - copy.start == limit;
                workbook.insertSheet(n++, copy);
            }
            // Close on the last copy worksheet
            workbook.getSheetAt(n - 1).shouldClose = true;
        } else {
            end = len;
        }
    }

    /**
     * 获取当前数组中有多少数据，数组中的数据是动态变化的，所以这是一个瞬时值
     *
     * @return 数组中数据大小
     */
    public int dataSize() {
        return data != null ? data.size() : 0;
    }

    /**
     * 除了实例化{@code ListSheet}工作表指定的数据外，导出过程中会使用{@code more}方法获取更多数据，
     * 直到返回{@code null}或空为止。这是一个独立于数据源的方法，它可以从任何数据源获取数据。
     * 由于此方法是无状态的，所以需要在自定义工作表中维护分页、请求参数和上下文信息。
     *
     * <p>每次获得的数据越多写入速度就越快, 更多的数据可以减少数据库查询或网络请求次数，
     * 但过多的数据会占用更多的内存，开发过程中需要在速度和内存消耗之间做权衡。
     * 建议批量不要太大，整个过程最大的压力在写磁盘操作上，所以批量太大对性能改善甚微，
     * 但最小批量最好要超过{@link #getRowBlockSize()}设置的大小也就是最小32</p>
     *
     * @return 数组，{@code null}和空数组表示结束
     */
    protected List<T> more() {
        if (dataSupplier != null) {
            int offset = left() + (rowBlock != null ? rowBlock.getTotal() : 0);
            if (copySheet) offset += copyCount * workbook.getSheetAt(id - 2).size();
            return dataSupplier.apply(offset, data != null && !data.isEmpty() ? data.get(data.size() - 1) : null);
        }
        return null;
    }

    /**
     * 添加“冻结”窗格，默认解析{@link FreezePanes}注解，支持自定义注解，
     * 只需包含冻结的行列信息即可
     *
     * @param clazz Class of &lt;T&gt;
     */
    protected void attachFreezePanes(Class<?> clazz) {
        // Annotation setting has lower priority than setting method
        if (getExtPropValue(Const.ExtendPropertyKey.FREEZE) != null) {
            return;
        }
        FreezePanes panes = clazz.getAnnotation(FreezePanes.class);
        if (panes == null) {
            return;
        }

        // Validity check
        if (panes.topRow() < 0 || panes.firstColumn() < 0) {
            throw new IllegalArgumentException("negative number occur.");
        }

        // Zero means unfreeze
        if ((panes.topRow() | panes.firstColumn()) == 0) {
            return;
        }

        // Put value into extend properties
        putExtProp(Const.ExtendPropertyKey.FREEZE, Panes.of(panes.topRow(), panes.firstColumn()));
    }

    /**
     * 查找其它可导出的{@code Method}方法，此方法将扩大查找范围，
     * 包含无参且仅有一个返回值的方法也做为查找对象，如果该方法标记有&#x40;ExcelColumn
     * 注解且无&#x40;IgnoreExport注解则添加到导出数组
     *
     * @param existsMethodMapper 已有的可导出{@code Method}映射，key为方法名
     * @param clazz              Class of &lt;T&gt;
     * @return 可导出的{@code Method}数组
     */
    protected Method[] filterOthersMethodsCanExport(Map<String, Method> existsMethodMapper, Class<?> clazz) {
        // Collect the method which has ExcelColumn annotation
        Method[] readMethods = null;
        try {
            Collection<Method> values = existsMethodMapper.values();
            readMethods = listReadMethods(clazz, method -> method.getAnnotation(ExcelColumn.class) != null
                && method.getAnnotation(IgnoreExport.class) == null && !values.contains(method));
        } catch (IntrospectionException e) {
            // Ignore
        }
        return readMethods;
    }

    /**
     * 强制导出
     *
     * <p>为了保证数据安全默认情况下Java Bean只导出标记有&#x40;ExcelColumn的字段和方法，
     * 但某些情况不方便修改实体此时可以使用强制导出功能将Bean中的全字段导出（标记有&#x40;IgnoreExport注解除外），
     * 此方法可能会造成数据泄漏风险，可参考{@link ExcelColumn}注解说明</p>
     *
     * @return 当前工作表
     */
    public Sheet forceExport() {
        this.forceExport = 1;
        return this;
    }

    /**
     * 取消强制导出，可以取消在工作表{@link Workbook}上设置的全局强制导出标记
     *
     * @return 当前工作表
     */
    public Sheet cancelForceExport() {
        this.forceExport = 2;
        return this;
    }

    /**
     * 返回“强制导出”标识
     *
     * @return 1: 强制导出 其余值：不强制
     */
    @Override
    public int getForceExport() {
        return forceExport;
    }

    /**
     * {@code ListSheet}独有的列对象，除了{@link Column}包含的信息外，它还保存当列对应的字段和方法，
     * 后续会通过这两个属性进行反射获取对象中的值，优先通过get方法获取，如果找不到get方法则直接
     * 使用{@link Field}获取值
     */
    public static class EntryColumn extends Column {
        /**
         * 当前列对应的get方法，这里不限get方法，无参且仅有一个返回值的方法均可以
         */
        public Method method;
        /**
         * 当前列对应的Bean字段
         */
        public Field field;

        public EntryColumn() {
            super();
        }

        public EntryColumn(String name) {
            super();
            this.name = name;
        }

        public EntryColumn(String name, Class<?> clazz) {
            super(name, clazz);
        }

        public EntryColumn(String name, String key) {
            super(name, key);
        }

        public EntryColumn(String name, String key, Class<?> clazz) {
            super(name, key, clazz);
        }

        public EntryColumn(String name, Class<?> clazz, ConversionProcessor processor) {
            super(name, clazz, processor);
        }

        public EntryColumn(String name, String key, ConversionProcessor processor) {
            super(name, key, processor);
        }

        public EntryColumn(String name, Class<?> clazz, boolean share) {
            super(name, clazz, share);
        }

        public EntryColumn(String name, String key, boolean share) {
            super(name, key, share);
        }

        public EntryColumn(String name, Class<?> clazz, ConversionProcessor processor, boolean share) {
            super(name, clazz, processor, share);
        }

        public EntryColumn(String name, String key, Class<?> clazz, ConversionProcessor processor) {
            super(name, key, clazz, processor);
        }

        public EntryColumn(String name, String key, ConversionProcessor processor, boolean share) {
            super(name, key, processor, share);
        }

        public EntryColumn(String name, Class<?> clazz, int cellStyle) {
            super(name, clazz, cellStyle);
        }

        public EntryColumn(String name, String key, int cellStyle) {
            super(name, key, cellStyle);
        }

        public EntryColumn(String name, Class<?> clazz, int cellStyle, boolean share) {
            super(name, clazz, cellStyle, share);
        }

        public EntryColumn(String name, String key, int cellStyle, boolean share) {
            super(name, key, cellStyle, share);
        }

        public EntryColumn(Column other) {
            super.from(other);
            if (other instanceof EntryColumn) {
                EntryColumn o = (EntryColumn) other;
                this.method = o.method;
                this.field = o.field;
            }
            if (other.next != null) {
                addSubColumn(new EntryColumn(other.next));
            }
        }

        /**
         * 获取当前列对应的get方法
         *
         * @return Method
         */
        public Method getMethod() {
            return method;
        }

        /**
         * 获取当前列对应的Bean字段
         *
         * @return Field
         */
        public Field getField() {
            return field;
        }
    }
}

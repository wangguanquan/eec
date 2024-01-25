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

import org.ttzero.excel.manager.Const;
import org.ttzero.excel.reader.Col;
import org.ttzero.excel.reader.Dimension;
import org.ttzero.excel.reader.ExcelReader;
import org.ttzero.excel.reader.FullSheet;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Path;
import java.util.Comparator;
import java.util.List;

/**
 * 模板工作表，它支持指定一个已有的Excel文件作为模板导出，{@code TemplateSheet}将复制
 * 模板工作表的样式并替换占位符，同时{@code TemplateSheet}也可以和其它{@code Sheet}共用，
 * 也可以有多个模板工作表，也就是最终的Excel可以包含多个模板源
 *
 * <p>创建模板工作表需要指定模板文件，它可以是本地文件也可是输入流{@code InputStream}，它同时支持{@code xls}
 * 和{@code xlsx}两种格式的模板，除模板文件外还需要指定Excel中的某个{@code Worksheet}，
 * 未指定工作表时默认以第一个工作表做为模板，{@code TemplateSheet}工作表导出时不受{@code ExcelColumn}注解限制，
 * 导出的数据范围由默认配置决定</p>
 *
 * <p>默认占位符为一对关闭的大括号{@code ‘${key}’}，</p>
 *
 * <p>考虑到模板工作表的复杂性暂时不支持切片查询数据，数据必须在初始化时设置，换句话说模板工作表只适用于少量数据</p>
 *
 * <blockquote><pre>
 * new Workbook("模板测试")
 *     .addSheet(new TemplateSheet(Paths.get("./template.xlsx")).setData(data)) // &lt;- 模板工作表
 *     .addSheet(new ListSheet<>()) // &lt;- 普通对象数组工作表
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
     * 源工作表
     */
    protected FullSheet sheet;

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
     * 实例化模板工作表并指定模板工作表索引，如果指定索引超过模板Excel中包含的工作表数量则抛异常
     *
     * @param templatePath 模板路径
     * @param originalSheetIndex 指定源工作表索引
     * @throws IOException 文件不存在或读取模板异常
     */
    public TemplateSheet(Path templatePath, int originalSheetIndex) throws IOException {
        this.reader = ExcelReader.read(templatePath);
        this.sheet = reader.sheet(originalSheetIndex).asFullSheet();
        if (sheet == null)
            throw new IOException("The specified index " + originalSheetIndex + " does not exist in template file.");
    }

    /**
     * 实例化模板工作表并指定模板工作表名，如果指定源工作表不存在则抛异常
     *
     * @param templatePath 模板路径
     * @param originalSheetName 指定源工作表名
     * @throws IOException 文件不存在或读取模板异常
     */
    public TemplateSheet(Path templatePath, String originalSheetName) throws IOException {
        this.reader = ExcelReader.read(templatePath);
        this.sheet = reader.sheet(originalSheetName).asFullSheet();
        if (sheet == null)
            throw new IOException("The specified sheet [" + originalSheetName + "] does not exist in template file.");
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
     * 实例化模板工作表并指定模板工作表索引，如果指定索引超过模板Excel中包含的工作表数量则抛异常
     *
     * @param templateStream 模板输入流
     * @param originalSheetIndex 指定源工作表索引
     * @throws IOException 读取模板异常
     */
    public TemplateSheet(InputStream templateStream, int originalSheetIndex) throws IOException {
        this.reader = ExcelReader.read(templateStream);
        this.sheet = reader.sheet(originalSheetIndex).asFullSheet();
        if (sheet == null)
            throw new IOException("The specified index " + originalSheetIndex + " does not exist in template file.");
    }

    /**
     * 实例化模板工作表并指定模板工作表名，如果指定源工作表不存在则抛异常
     *
     * @param templateStream 模板输入流
     * @param originalSheetName 指定源工作表名
     * @throws IOException 读取模板异常
     */
    public TemplateSheet(InputStream templateStream, String originalSheetName) throws IOException {
        this.reader = ExcelReader.read(templateStream);
        this.sheet = reader.sheet(originalSheetName).asFullSheet();
        if (sheet == null)
            throw new IOException("The specified sheet [" + originalSheetName + "] does not exist in template file.");
    }

    @Override
    protected Column[] getHeaderColumns() {
        if (!headerReady) {
            // 解析模板工作表并复制信息到当前工作表中
            int size = init();
            if (size <= 0) {
                columns = new Column[0];
            }
        }
        return columns;
    }

    protected int init() {
        // 冻结,直接复制不需要计算移动
        Panes panes = sheet.getFreezePanes();
        if (panes != null) putExtProp(Const.ExtendPropertyKey.FREEZE, panes);

        // TODO 合并（较为复杂不能简单复制，需要计算中间插入或扣除的行）
        List<Dimension> mergeCells = sheet.getMergeCells();
        if (panes != null) putExtProp(Const.ExtendPropertyKey.MERGE_CELLS, mergeCells);

        // 过滤
        Dimension autoFilter = sheet.getFilter();
        if (autoFilter != null) putExtProp(Const.ExtendPropertyKey.AUTO_FILTER, autoFilter);

        // 是否显示网格线
        this.showGridLines = sheet.showGridLines();

        // 获取列属性
        List<Col> cols = sheet.getCols();
        cols.sort(Comparator.comparingInt(a -> a.max));
        // 创建列
        int len = cols.get(cols.size() - 1).max, i = 0;
        columns = new Column[len];
        for (Col col : cols) {
            for (int a = col.min; a <= col.max; a++) {
                Column c = new Column();
                c.width = col.width;
                c.colIndex = a - 1;
                if (col.hidden) c.hide();
                columns[i++] = c;
            }
        }
        // 忽略表头输出
        super.ignoreHeader();

        return len;
    }

    @Override
    protected void resetBlockData() {

    }

    @Override
    public void close() throws IOException {
        super.close();
        // 释放模板流
        if (reader != null) reader.close();
    }
}

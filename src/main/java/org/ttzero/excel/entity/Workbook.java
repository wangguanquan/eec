/*
 * Copyright (c) 2017, guanquan.wang@hotmail.com All Rights Reserved.
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

import org.ttzero.excel.entity.csv.CSVWorkbookWriter;
import org.ttzero.excel.entity.e7.ContentType;
import org.ttzero.excel.entity.e7.XMLWorkbookWriter;
import org.ttzero.excel.entity.style.Fill;
import org.ttzero.excel.entity.style.PatternType;
import org.ttzero.excel.entity.style.Styles;
import org.ttzero.excel.manager.docProps.Core;
import org.ttzero.excel.manager.docProps.CustomProperties;
import org.ttzero.excel.util.StringUtil;

import java.awt.Color;
import java.io.File;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.charset.Charset;
import java.nio.file.Path;
import java.util.Arrays;
import java.util.Map;
import java.util.function.BiConsumer;
import java.util.zip.Deflater;


/**
 * 一个{@code Workbook}工作薄实例即表示一个Excel文件，它包含一个或多个{@link Sheet}工作表，
 * Workbook收集全局属性，如文档属性、样式，字符串共享区等。
 *
 * <p>在导出Excel文件时需要遵循以下三个步骤：</p>
 * <ol>
 *     <li>设置文件属性（非必须）</li>
 *     <li>调用{@link #addSheet}添加Worksheet工作表（必须）</li>
 *     <li>调用{@link #writeTo}方法来执行写入操作（必须）</li>
 * </ol>
 * <p>当前仅支持xlsx(默认）和csv格式输出，保存为csv格式须在调用writeTo之前调用{@link #saveAsCSV()}方法，
 * 如果当前Workbook包含多个Worksheet则会生成多个csv文件并将多个文件压缩为zip格式。</p>
 *
 * <p>常用的属性包含{@link #setCreator}设置作者, {@link #setCompany}设置公司名,
 * {@link #setAutoSize}设置自适应列宽和{@link #setZebraLine}设置斑马线，
 * 前两种需要打开文件详细属性查看，后两种起美化作用有利于阅读，后两个属性可以设置到Workbook或各Worksheet中，
 * 如果设置到Workbook则会应用于所有Worksheet，当Workbook和Worksheet均设置了同一属性则Worksheet优先。</p>
 *
 * <p>{@link #writeTo}方法是一个终止符它将执行实际的写操作，所以需要将该方法放在所有语句之后，
 * 任何放置在该方法之后的指令将会被忽略。</p>
 *
 * <p>本工具将数据源和输出协议分开设计，工作表Sheet为数据源，{@link IWorkbookWriter}和{@link IWorksheetWriter}
 * 为输出协议，实现不同的输出协议即实现不同格式化输出，已实现的{@link org.ttzero.excel.entity.e7.XMLWorksheetWriter}和
 * {@link org.ttzero.excel.entity.csv.CSVWorksheetWriter}就是xlsx和csv格式的输出协议实现。
 * 工作薄Workbook对应的输出协议为{@link IWorkbookWriter}，它负责协调所有部件输出并将所有零散的文件组装为OpenXml格式。</p>
 *
 * <p>一个典型的导出示例：
 * <pre>
 * new Workbook("双11销量统计")
 *     // 设置作者
 *     .setCreator("作者")
 *     // 设置自适应列宽
 *     .setAutoSize(true)
 *     // 添加一个名为"总销量排行"的Worksheet
 *     .addSheet(new ListSheet&lt;Item&gt;("总销量排行")
 *         .setData(new ArrayList&lt;&gt;())) // &lt;- 这里替换为实际数据
 *     // 添加一个名为"单品销量排行"的Worksheet
 *     .addSheet(new ListMapSheet&lt;&gt;("单品销量排行")
 *         .setData(new ArrayList&lt;&gt;())) // &lt;- 这里替换为实际数据
 *     // 指定输出路径 '/tmp/"双11销量统计".xlsx'
 *     .writeTo(Paths.get("/tmp/"));</pre>
 *
 * <p>参考文档:</p>
 * <p><a href="https://poi.apache.org">POI</a></p>
 * <p><a href="https://msdn.microsoft.com/library">Office 365</a></p>
 * <p><a href="https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet(v=office.14).aspx#">DocumentFormat.OpenXml.Spreadsheet Namespace</a></p>
 * <p><a href="https://docs.microsoft.com/zh-cn/previous-versions/office/office-12/ms406049(v=office.12)">介绍 Microsoft Office (2007) Open XML 文件格式</a></p>
 *
 * @author guanquan.wang on 2017/9/26.
 */
public class Workbook implements Storable {
    /**
     * 工作薄名
     */
    private String name;
    /**
     * 工作表数组，按数组顺序输出
     */
    private Sheet[] sheets;
    /**
     * 水印
     */
    private Watermark watermark;
    /**
     * 记录工作表几数
     */
    private int size;
    /**
     * 全局自适应列宽标识
     */
    private boolean autoSize;
    /**
     * 作者，未指定时将默认取当前系统登录名
     */
    private String creator;
    /**
     * 工作薄属性
     */
    private Core core;
    /**
     * 公司名
     */
    private String company;
    /**
     * 全局斑马线
     */
    private Fill zebraFill;
    /**
     * 导出进度监控器
     */
    private BiConsumer<Sheet, Integer> progressConsumer;
    /**
     * 全局字符串共享区
     */
    private SharedStrings sst;
    /**
     * 全局样式
     */
    private Styles styles;
    /**
     * WorkbookWriter输出协议，输出协议影响最终的文件格式
     */
    private IWorkbookWriter workbookWriter;
    /**
     * 强制导出，绕过安全限制导出全字段
     */
    private int forceExport;
    /**
     * 全局ContentType
     */
    private final ContentType contentType;
    /**
     * 全局Drawing记数器
     */
    private int drawingCounter;
    /**
     * 全局多媒体记数器（当前仅支持图片）
     */
    private int mediaCounter;
    /**
     * 自定义属性
     */
    private CustomProperties customProperties;
    /**
     * 压缩等级 {@code 0-9}，数字越小压缩效果越好耗时越长
     */
    private int compressionLevel = 5;

    /**
     * 创建一个未命名工作薄
     *
     * <p>如果writeTo方法指定的File或Path为文件夹时，未命名工作薄将会以'新建文件'作为文件名</p>
     */
    public Workbook() {
        this(null);
    }

    /**
     * 创建一个工作薄并指定名称
     *
     * @param name 工作薄名
     */
    public Workbook(String name) {
        this(name, null);
    }

    /**
     * 创建一个工作薄并指定名称和作者
     *
     * @param name    工作薄名
     * @param creator 作者
     */
    public Workbook(String name, String creator) {
        this.name = name;
        this.creator = creator;
        sheets = new Sheet[3]; // Create three worksheet
        contentType = new ContentType();
    }

    /**
     * 获取当前工作薄名称
     *
     * @return 工作薄名称
     */
    public String getName() {
        return name;
    }

    /**
     * 设置工作薄名称，如果writeTo方法指定的Path或File为文件夹时该名称将作为最终文件名
     *
     * @param name 工作薄名，长度最好不超过255个字符
     * @return 当前工作薄
     */
    public Workbook setName(String name) {
        this.name = name;
        return this;
    }

    /**
     * 获取当前工作薄作者
     *
     * @return 作者
     */
    public String getCreator() {
        return creator;
    }

    /**
     * 获取当前工作薄公司名
     *
     * @return 公司名
     */
    public String getCompany() {
        return company;
    }

    /**
     * 获取当前工作薄包含的工作表个数
     *
     * @return 工作表个数
     */
    public int getSize() {
        return size;
    }

    /**
     * 获取文档属性，包含主题，关键词，分类等信息
     *
     * @return 文档属性
     */
    public Core getCore() {
        return core;
    }

    /**
     * 设置文档属性，包含主题，关键词，分类等信息
     *
     * @param core 文档属性
     * @return 当前工作薄
     */
    public Workbook setCore(Core core) {
        this.core = core;
        return this;
    }

    /**
     * 获取全局字符串共享区，此共享区独立于Worksheet，所有worksheet共享
     *
     * @return 全局字符串共享区{@link SharedStrings}
     */
    public SharedStrings getSharedStrings() {
        // CSV do not need SharedStringTable
        if (!(workbookWriter instanceof CSVWorkbookWriter) && sst == null)
            sst = new SharedStrings();
        return sst;
    }

    /**
     * 获取所有{@link Sheet}集合
     *
     * <p>注意：返回的对象是一个浅拷贝对其做任何修改将影响最终效果</p>
     *
     * @return {@link Sheet}集合
     */
    public final Sheet[] getSheets() {
        return Arrays.copyOf(sheets, size);
    }

    /**
     * 获取水印{@link Watermark}
     *
     * @return 水印
     */
    public Watermark getWatermark() {
        return watermark;
    }

    /**
     * 设置水印{@link Watermark}，可以使用{@link Watermark#of}静态方法创建
     *
     * @param watermark 水印
     * @return 当前工作薄
     */
    public Workbook setWatermark(Watermark watermark) {
        this.watermark = watermark;
        return this;
    }

    /**
     * 获取水印{@link Watermark}
     *
     * @return 水印
     * @deprecated 重命名为 {@link #getWatermark()}
     */
    @Deprecated
    public Watermark getWaterMark() {
        return getWatermark();
    }

    /**
     * 设置水印{@link Watermark}，可以使用{@link Watermark#of}静态方法创建
     *
     * @param watermark 水印
     * @return 当前工作薄
     * @deprecated 重命名为 {@link #setWatermark(Watermark)}
     */
    @Deprecated
    public Workbook setWaterMark(Watermark watermark) {
        return setWatermark(watermark);
    }

    /**
     * 设置全局自适应列宽
     *
     * @param autoSize true: 自适应宽度，false：固定宽度（默认）
     * @return 当前工作薄
     */
    public Workbook setAutoSize(boolean autoSize) {
        this.autoSize = autoSize;
        return this;
    }

    /**
     * 获取当前工作薄是否为自适应宽度
     *
     * @return true: 自适应宽度，false：固定宽度
     */
    public boolean isAutoSize() {
        return autoSize;
    }

    /**
     * 强制导出
     * 
     * <p>注意：设置此标记后将无视安全规则导出Java对象中的所有字段，请根据实际情况谨慎使用</p>
     *
     * @return 当前工作薄
     */
    public Workbook forceExport() {
        this.forceExport = 1;
        return this;
    }

    /**
     * 获取当前工作薄是否为“强制导出”
     *
     * @return 强制导出时返回1，其它情况返回0
     */
    public int getForceExport() {
        return forceExport;
    }

    /**
     * 获取全局样式{@link Styles}
     *
     * @return 全局样式
     */
    public Styles getStyles() {
        // CSV do not need Styles
        if (styles == null && !(workbookWriter instanceof CSVWorkbookWriter))
            styles = Styles.create();
        return styles;
    }

    /**
     * 设置全局样式{@link Styles}
     *
     * @param styles 定制化全局样式
     * @return 当前工作薄
     */
    public Workbook setStyles(Styles styles) {
        this.styles = styles;
        return this;
    }

    /**
     * 设置作者
     *
     * <p>设置作者后可以通过查看文件属性来查看作者</p>
     *
     * @param creator 作者
     * @return 当前工作薄
     */
    public Workbook setCreator(String creator) {
        this.creator = creator;
        return this;
    }

    /**
     * 设置公司名，建议控制在64个字符以内
     *
     * @param company 公司名
     * @return 当前工作薄
     */
    public Workbook setCompany(String company) {
        this.company = company;
        return this;
    }

    /**
     * 设置斑马线背景，斑马线是由相同间隔的背景色造成的视觉效果，有助于从视觉上区分每行数据，
     * 但刺眼的背景色可能造成相反的效果，设置之前最好在Office中提前预览效果
     *
     * @param fill 背景样式{@link Fill}
     * @return 当前工作薄
     */
    public Workbook setZebraLine(Fill fill) {
        this.zebraFill = fill;
        return this;
    }

    /**
     * 取消斑马线
     *
     * @return 当前工作薄
     */
    public Workbook cancelZebraLine() {
        this.zebraFill = null;
        return this;
    }

    /**
     * 指定以默认斑马线输出，默认背景颜色为{@code #EFF5EB}
     *
     * @return 当前工作薄
     */
    public Workbook defaultZebraLine() {
        return setZebraLine(new Fill(PatternType.solid, new Color(233, 234, 236)));
    }

    /**
     * 获取斑马线背景样式
     *
     * @return 斑马线背景 {@link Fill}样式
     */
    public Fill getZebraFill() {
        return zebraFill;
    }

    /**
     * 判断当前工作薄是否设置了全局斑马线背景
     *
     * @return true: 有全局斑马线
     */
    public boolean hasZebraFill() {
        return zebraFill != null && zebraFill.getPatternType() != PatternType.none;
    }

    /**
     * 获取工作薄输出协议{@link IWorkbookWriter}
     *
     * @return 工作薄输出协议
     */
    public IWorkbookWriter getWorkbookWriter() {
        if (workbookWriter == null)
            workbookWriter = new XMLWorkbookWriter(this);
        return workbookWriter;
    }

    /**
     * 添加一个工作表{@link Sheet}，新添加的工作表总是排在队列最后，
     * 可以使用{@link #insertSheet}插入到指定位置
     *
     * @param sheet 工作表
     * @return 当前工作薄
     */
    public Workbook addSheet(Sheet sheet) {
        ensureCapacityInternal();
        sheet.setWorkbook(this);
        sheets[size++] = sheet;
        return this;
    }

    /**
     * 在指定下标插入一个工作表{@link Sheet}
     *
     * @param index 指定工作表插入的位置（从0开始）
     * @param sheet 待插入的工作表
     * @return 当前工作薄
     */
    public Workbook insertSheet(int index, Sheet sheet) {
        ensureCapacityInternal();
        int _size = size;
        if (sheets[index] != null) {
            for (; _size > index; _size--) {
                sheets[_size] = sheets[_size - 1];
                sheets[_size].setId(sheets[_size].getId() + 1);
            }
        }
        sheets[index] = sheet;
        sheet.setId(index + 1);
        sheet.setWorkbook(this);
        size++;
        return this;
    }

    /**
     * 移除指定位置的工作表{@link Sheet}
     *
     * @param index 待移除的工作表下标（从0开始）
     * @return 当前工作薄
     */
    public Workbook remove(int index) {
        if (index < 0 || index >= size) {
            return this;
        }
        if (index == size - 1) {
            sheets[index] = null;
        } else {
            for (; index < size - 1; index++) {
                sheets[index] = sheets[index + 1];
                sheets[index].setId(sheets[index].getId() - 1);
            }
        }
        size--;
        return this;
    }

    /**
     * 获取指定位置的工作表
     *
     * <p>如果使用{@link #insertSheet}方法插入了一个较大的下标，调用此方法可能返回null值。
     * 例如在下标为100的位置插入了一个工作表，获取第90位的工作薄将返回一个null值。</p>
     *
     * @param index 工作表在队列中的位置（从0开始）
     * @return 指定位置的工作表 {@link Sheet}
     * @throws IndexOutOfBoundsException 如果下标为负数或者超过工作薄队列长度
     */
    public Sheet getSheetAt(int index) {
        if (index < 0 || index >= size)
            throw new IndexOutOfBoundsException("Index: " + index + ", Size: " + size);
        return sheets[index];
    }

    /**
     * 返回指定名称的工作表{@link Sheet}
     *
     * <p>注意：只能查找那些在创建时设置了名称的工作表</p>
     *
     * @param sheetName 待查找的工作表名称
     * @return 按工作表名称查询，未找到时返回{code null}
     */
    public Sheet getSheet(String sheetName) {
        if (StringUtil.isEmpty(sheetName)) return null;
        for (Sheet sheet : sheets) {
            if (sheetName.equals(sheet.getName())) {
                return sheet;
            }
        }
        return null;
    }

    /**
     * 添加一个进度监听器，可以在较大导出时展示进度
     *
     * <pre>
     * new Workbook().onProgress((sheet, row) -&gt; {
     *     System.out.println(sheet + " write " + row + " rows");
     * })</pre>
     *
     * @param progressConsumer 进度监听器
     * @return 当前工作薄
     */
    public Workbook onProgress(BiConsumer<Sheet, Integer> progressConsumer) {
        this.progressConsumer = progressConsumer;
        return this;
    }

    /**
     * 获取进度监听器
     *
     * @return 如果设置了监听器就返回，未设置时返回null
     */
    public BiConsumer<Sheet, Integer> getProgressConsumer() {
        return progressConsumer;
    }

//    /**
//     * Save as excel97~2003
//     * <p>
//     * You mast add eec-e3-support.jar into class path to support excel97~2003
//     *
//     * @return 当前工作薄
//     * @throws OperationNotSupportedException if eec-e3-support not import into class path
//     */
//    public Workbook saveAsExcel2003() throws OperationNotSupportedException {
//        try {
//            // Create Styles and SharedStringTable
//            Class<?> clazz = Class.forName("org.ttzero.excel.entity.e3.BIFF8WorkbookWriter");
//            Constructor<?> constructor = clazz.getDeclaredConstructor(this.getClass());
//            workbookWriter = (IWorkbookWriter) constructor.newInstance(this);
//        } catch (Exception e) {
//            throw new OperationNotSupportedException("Excel97-2003 Not support now.");
//        }
//        return this;
//    }

    /**
     * 另存为Comma-Separated Values格式，默认使用','逗号分隔
     *
     * @return 当前工作薄
     */
    public Workbook saveAsCSV() {
        workbookWriter = new CSVWorkbookWriter(this);
        return this;
    }

    /**
     * 另存为Comma-Separated Values格式并保存BOM，默认使用','逗号分隔
     *
     * @return 当前工作薄
     */
    public Workbook saveAsCSVWithBom() {
        workbookWriter = new CSVWorkbookWriter(this, true);
        return this;
    }

    /**
     * 以指定字符集保存为Comma-Separated Values格式，默认使用','逗号分隔
     *
     * @param charset 指定输出字符集
     * @return 当前工作薄
     */
    public Workbook saveAsCSV(Charset charset) {
        workbookWriter = new CSVWorkbookWriter(this).setCharset(charset);
        return this;
    }

    /**
     * 以指定字符集保存为Comma-Separated Values格式并保存BOM，默认使用','逗号分隔
     *
     * @param charset 指定输出字符集
     * @return 当前工作薄
     */
    public Workbook saveAsCSVWithBom(Charset charset) {
        workbookWriter = new CSVWorkbookWriter(this, true).setCharset(charset);
        return this;
    }

    /**
     * 确认边距并在越界时自动扩容
     */
    private void ensureCapacityInternal() {
        if (size >= sheets.length) {
            sheets = Arrays.copyOf(sheets, size + 1);
        }
    }

    //////////////////////////Print Out/////////////////////////////

    /**
     * 指定输出路径，Path可以是文件夹或者文件
     *
     * <p>如果Path为文件夹时将在该文件夹下生成名为{@link #getName()} + {@link IWorkbookWriter#getSuffix()}的文件，
     * 文件后缀随输出协议变动。如果存在相同的文件名则会在文件名后面添加'(n)'，n为自增的数字，例已存在"abc.xlsx"，
     * 再次导出时将保存为"abc(1).xlsx"。如果Path为明确的文件绝对路径，那将保存在Path的绝对路径下，已存在相同文件时会覆盖原文件，
     * 需要注意覆盖失败的情况</p>
     *
     * @param path Excel保存位置
     * @throws IOException I/O操作异常
     */
    @Override
    public void writeTo(Path path) throws IOException {
        checkAndInitWriter();
        try {
            workbookWriter.writeTo(path);
        } finally {
            workbookWriter.close();
        }
    }

    /**
     * 导出到{@link OutputStream}流，适用于小文件Excel直接导出的场景
     *
     * <pre>
     * public void export(HttpServletResponse response) throws IOException {
     *     String fileName = java.net.URLEncoder.encode("abc.xlsx", "UTF-8");
     *     response.setHeader(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=\""
     *         + fileName + "\"; filename*=utf-8''" + fileName);
     *     new Workbook()
     *         .addSheet(new ListSheet&lt;Item&gt;("总销量排行", new ArrayList&lt;&gt;()))
     *         // 直接写到Response
     *         .writeTo(response.getOutputStream());
     * }</pre>
     *
     * @param os 输出流
     * @throws IOException         I/O操作异常
     * @throws ExcelWriteException 其它运行时异常
     */
    public void writeTo(OutputStream os) throws IOException, ExcelWriteException {
        checkAndInitWriter();
        try {
            workbookWriter.writeTo(os);
        } finally {
            workbookWriter.close();
        }
    }

    /**
     * 指定输出路径，File可以是文件夹或者文件
     *
     * <p>如果File为文件夹时将在该文件夹下生成名为{@link #getName()} + {@link IWorkbookWriter#getSuffix()}的文件，
     * 文件后缀随输出协议变动。如果存在相同的文件名则会在文件名后面添加'(n)'，n为自增的数字，例已存在"abc.xlsx"，
     * 再次导出时将保存为"abc(1).xlsx"。如果File为明确的文件绝对路径，那将保存在File的绝对路径下，已存在相同文件时会覆盖原文件，
     * 需要注意覆盖失败的情况</p>
     *
     * @param file                 Excel保存位置
     * @throws IOException         I/O操作异常
     */
    public void writeTo(File file) throws IOException {
        writeTo(file.toPath());
    }

    /**
     * 设置自定义工作薄输出协议
     *
     * @param workbookWriter 自定义工作薄{@link IWorkbookWriter}协议
     * @return 当前工作薄
     */
    public Workbook setWorkbookWriter(IWorkbookWriter workbookWriter) {
        this.workbookWriter = workbookWriter;
        this.workbookWriter.setWorkbook(this);
        return this;
    }

    /**
     * 初始化，创建全局样式和字符串共享区
     * @deprecated 不需要主动初始化，后续将删除
     */
    @Deprecated
    protected void init() {
        // 创建全局字符串共享区
        if (sst == null) {
            sst = new SharedStrings();
        }
        // 创建全局样式
        if (styles == null) {
            styles = Styles.create();
        }
    }

    /**
     * 检查并创建工作薄协议{@link IWorkbookWriter}
     */
    protected void checkAndInitWriter() {
        if (workbookWriter == null) {
//            // 初始化
//            init();
            workbookWriter = new XMLWorkbookWriter(this);
        }
    }

    /**
     * 添加资源类型，导出图片时按照图片格式添加不同的资源类型，一般情况下开发者不需要关心
     *
     * @param type 资源类型{@link ContentType.Type}
     * @return 当前工作薄
     */
    public Workbook addContentType(ContentType.Type type) {
        contentType.add(type);
        return this;
    }

    /**
     * 添加ContentType关系，一般情况下开发者不需要关心
     *
     * @param rel {@link Relationship}关系
     * @return 当前工作薄
     */
    public Workbook addContentTypeRel(Relationship rel) {
        contentType.addRel(rel);
        return this;
    }

    /**
     * 获取全局的资源类型，一般情况下开发者不需要关心
     *
     * @return 资源类型{@link ContentType}对象
     */
    public ContentType getContentType() {
        return contentType;
    }

    /**
     * 图片记数器自增
     *
     * @return 图片记数器
     */
    public int incrementDrawingCounter() {
        return ++drawingCounter;
    }

    /**
     * 获取当前工作薄包含多少张图片
     *
     * @return 图片数量
     */
    public int getDrawingCounter() {
        return drawingCounter;
    }

    /**
     * 媒体记数器，一般情况下media与worksheet对应
     *
     * @return 媒体记数器
     */
    public int incrementMediaCounter() {
        return ++mediaCounter;
    }

    /**
     * 获取当前工作薄含有多媒体的工作表个数
     *
     * @return 含有多媒体的工作表个数
     */
    public int getMediaCounter() {
        return mediaCounter;
    }

    /**
     * 添加自定义属性，自定义属性可以从"信息"-&gt;"属性"-&gt;"自定义属性"查看
     *
     * <p>注意：只支持{@code "文本"}、{@code "数字"}、{@code "日期"}以及{@code "布尔值"}，其它数据类型将使用{@code toString}强转换为文本</p>
     *
     * @param key 属性名，不超过{@code 256}个字符
     * @param value 属性值，
     * @return 当前工作表
     */
    public Workbook putCustomProperty(String key, Object value) {
        if (customProperties == null) customProperties = new CustomProperties();
        customProperties.put(key, value);
        return this;
    }

    /**
     * 添加自定义属性，自定义属性可以从"信息"-&gt;"属性"-&gt;"自定义属性"查看
     *
     * <p>注意：只支持{@code "文本"}、{@code "数字"}、{@code "日期"}以及{@code "布尔值"}，其它数据类型将使用{@code toString}强转换为文本</p>
     *
     * @param properties 批量属性
     * @return 当前工作表
     */
    public Workbook putCustomProperties(Map<String, Object> properties) {
        if (customProperties == null) customProperties = new CustomProperties();
        customProperties.putAll(properties);
        return this;
    }

    /**
     * 删除自定义属性
     *
     * @param key 指定属性名
     * @return 如果属性存在则返回属性值否则返回 {@code null}
     */
    public Object removeCustomProperty(String key) {
        return customProperties != null ? customProperties.remove(key) : null;
    }

    /**
     * 获取自定义属性类
     *
     * @return {@code Custom}自定义属性类
     */
    public CustomProperties getCustomProperties() {
        return customProperties;
    }

    /**
     * 文档保护-标记只读
     *
     * @return 当前工作表
     */
    public Workbook markAsReadOnly() {
        if (customProperties == null) customProperties = new CustomProperties();
        customProperties.markAsReadOnly();
        return this;
    }

    /**
     * 将压缩等级设为{@code 1}以获取更快的速度
     *
     * @return 当前工作表
     */
    public Workbook bestSpeed() {
        this.compressionLevel = Deflater.BEST_SPEED;
        return this;
    }

    /**
     * 获取压缩等级
     *
     * @return 压缩等级
     */
    public int getCompressionLevel() {
        return Math.min(Math.max(compressionLevel, 0), 9);
    }
}

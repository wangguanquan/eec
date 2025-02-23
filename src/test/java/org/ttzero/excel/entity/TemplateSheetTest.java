/*
 * Copyright (c) 2017-2024, guanquan.wang@yandex.com All Rights Reserved.
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

import org.junit.Ignore;
import org.junit.Test;
import org.ttzero.excel.entity.e7.XMLCellValueAndStyle;
import org.ttzero.excel.entity.e7.XMLWorkbookWriter;
import org.ttzero.excel.entity.style.Fill;
import org.ttzero.excel.entity.style.Font;
import org.ttzero.excel.entity.style.Horizontals;
import org.ttzero.excel.entity.style.NumFmt;
import org.ttzero.excel.entity.style.Styles;
import org.ttzero.excel.manager.Const;
import org.ttzero.excel.manager.ExcelType;
import org.ttzero.excel.reader.Cell;
import org.ttzero.excel.reader.Dimension;
import org.ttzero.excel.reader.ExcelReader;
import org.ttzero.excel.reader.FullSheet;
import org.ttzero.excel.reader.Row;
import org.ttzero.excel.util.FileUtil;
import org.ttzero.excel.util.StringUtil;
import org.ttzero.excel.util.ZipUtil;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.sql.Timestamp;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.function.Supplier;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertTrue;
import static org.ttzero.excel.reader.ExcelReaderTest.testResourceRoot;

/**
 * @author guanquan.wang at 2024-01-25 09:57
 */
public class TemplateSheetTest extends WorkbookTest {

    @Test public void testMultiTemplates() throws IOException {
        final String fileName = "multi template sheets.xlsx";
        new Workbook()
            .addSheet(new TemplateSheet("模板 1.xlsx", testResourceRoot().resolve("1.xlsx"))) // <- 模板工作表
            .addSheet(new ListSheet<>("普通工作表", ListObjectSheetTest.Item.randomTestData())) // <- 普通工作表
            .addSheet(new TemplateSheet("模板 fracture merged.xlsx", testResourceRoot().resolve("fracture merged.xlsx"))) // <- 模板工作表
            .addSheet(new TemplateSheet("复制空白工作表", testResourceRoot().resolve("#81.xlsx"), "Sheet2")) // 空白工作表模板
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            assertEquals(reader.getSheetCount(), 4);
            // TODO 判断每个工作表的内容和样式
        }
    }

    @Test public void testAllTemplateSheets() throws IOException {
        final String fileName = "all template sheets.xlsx";
        Workbook workbook = new Workbook();
        File[] files = testResourceRoot().toFile().listFiles();
        if (files != null) {
            for (File file : files) {
                if (ExcelReader.getType(file.toPath()) == ExcelType.XLSX) {
                    try (ExcelReader reader = ExcelReader.read(file.toPath())) {
                        org.ttzero.excel.reader.Sheet[] sheets = reader.all();
                        // 这里设置占位符前缀为[#@^!]是为了全量复制数据用
                        for (org.ttzero.excel.reader.Sheet sheet : sheets) {
                            workbook.addSheet(new TemplateSheet(file.getName() + "$" + sheet.getName(), file.toPath(), sheet.getName()).setPrefix("#@^!"));
                        }
                    }
                }
            }
        }
        workbook.writeTo(getOutputTestPath().resolve(fileName));
    }

    @Test public void testSimpleTemplate() throws IOException {
        final String fileName = "simple template.xlsx";
        List<GameEntry> expectList = GameEntry.random();
        new Workbook()
            .addSheet(new TemplateSheet(testResourceRoot().resolve("template2.xlsx"), "简单模板").setData(expectList))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            FullSheet sheet = reader.sheet(0).asFullSheet();
            Dimension autoFilter = sheet.getFilter();
            assertEquals(autoFilter, Dimension.of("A1:F1"));
            Iterator<org.ttzero.excel.reader.Row> iter = reader.sheet(0).asFullSheet().iterator();
            Styles styles = reader.getStyles();
            assertTrue(iter.hasNext());
            org.ttzero.excel.reader.Row row0 = iter.next();
            for (int i = row0.getFirstColumnIndex(), len = row0.getLastColumnIndex(); i < len; i++) {
                Cell cell = row0.getCell(i);
                int style = row0.getCellStyle(cell);
                Fill fill = styles.getFill(style);
                assertEquals(fill.getFgColor(), new java.awt.Color(112, 173, 71));
            }
            List<GameEntry> list = new ArrayList<>();
            for (; iter.hasNext(); ) {
                org.ttzero.excel.reader.Row row = iter.next();
                GameEntry e = new GameEntry();
                e.channel = row.getInt(0);
                e.game = row.getString(1);
                e.account = row.getString(2);
                e.date = row.getDate(3);
                e.isAdult = row.getBoolean(4);
                e.vip = row.getString(5);
                list.add(e);
            }
            assertEquals(expectList.size(), list.size());
            for (int i = 0, len = expectList.size(); i < len; i++) {
                GameEntry expect = expectList.get(i), o = list.get(i);
                assertEquals(expect, o);
            }
        }
    }

    @Test public void testTemplate() throws IOException {
        Map<String, Object> map = new HashMap<>();
        map.put("name", author);
        map.put("score", random.nextInt(90) + 10);
        map.put("date", LocalDate.now().toString());
        map.put("desc", "暑假");

        new Workbook()
            .addSheet(new TemplateSheet(Files.newInputStream(testResourceRoot().resolve("template.xlsx")))
                .setData(map))
            .writeTo(defaultTestPath.resolve("fill inputstream template with map.xlsx"));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("fill inputstream template with map.xlsx"))) {
            for (Iterator<Row> it = reader.sheet(0).iterator(); it.hasNext(); ) {
                Row row = it.next();
                switch (row.getRowNum()) {
                    case 1:
                        assertEquals("通知书", row.getString(0).trim());
                        break;
                    case 3:
                        assertEquals((map.get("name") + " 同学，在本次期末考试的成绩是 " + map.get("score")+ "，希望"), row.getString(1).trim());
                        break;
                    case 4:
                        assertEquals(("下学期继续努力，祝你有一个愉快的" + map.get("desc") + "。"), row.getString(0).trim());
                        break;
                    case 23:
                        assertEquals(map.get("date"), row.getString(0).trim());
                        break;
                    default:
                        assertTrue(row.isBlank());
                }
            }
        }
    }

    @Test public void testFillObject() throws IOException {
        final String fileName = "fill object.xlsx";
        YzEntity yzEntity = YzEntity.mock();
        YzOrderEntity yzOrderEntity = YzOrderEntity.mock();
        YzSummary yzSummary = YzSummary.mock();
        new Workbook()
            .addSheet(new TemplateSheet(testResourceRoot().resolve("template2.xlsx"), "混合命名空间")
                .setData(yzEntity)
                .setData("YzEntity", yzOrderEntity)
                .setData("summary", yzSummary)
            ).writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            FullSheet sheet = reader.sheet(0).asFullSheet();
            Styles styles = reader.getStyles();
            Iterator<org.ttzero.excel.reader.Row> iter = sheet.iterator();
            // 第一行
            assertTrue(iter.hasNext());
            org.ttzero.excel.reader.Row row = iter.next();
            assertEquals(row.getString(0), yzEntity.gsName + "精品采购订单");
            Font font = styles.getFont(row.getCellStyle(0));
            assertEquals(font.getName(), "宋体");
            assertEquals(font.getSize(), 20);
            assertTrue(font.isBold());

            // 第二行
            assertTrue(iter.hasNext());
            row = iter.next();
            assertEquals(yzEntity.gysName, row.getString(3));
            assertEquals(yzEntity.orderNo, row.getString(11));

            // 第三行
            assertTrue(iter.hasNext());
            row = iter.next();
            assertEquals(yzEntity.gsName, row.getString(3));
            assertEquals(yzEntity.orderStatus, row.getString(11));

            // 第四行
            assertTrue(iter.hasNext());
            row = iter.next();
            assertEquals(yzEntity.jsName, row.getString(3));
            assertTrue(yzEntity.cgDate.getTime() - row.getDate(11).getTime() < 1000); // 导出时Excel的日期丢失了毫秒值

            // 第五行
            assertTrue(iter.hasNext());
            assertTrue(iter.next().isBlank()); // 空行

            // 第六行
            assertTrue(iter.hasNext());
            row = iter.next();
            assertEquals(row.getFirstColumnIndex(), 0);
            final String[] titles = {"序号", "精品代码", null, "精品名称", null, null, "数量", "不含税单价", "不含税金额", "税率", "含税单价", "含税金额", "备注"};
            for (int i = 0, len = Math.min(row.getLastColumnIndex(), titles.length); i < len; i++) {
                Cell cell = row.getCell(i);
                assertEquals(titles[i], row.getString(cell));
                if (titles[i] != null) {
                    int style = row.getCellStyle(cell);
                    font = styles.getFont(style);
                    assertTrue(font.isBold());
                    assertEquals(font.getName(), "宋体");
                    assertEquals(font.getSize(), 11);
                    assertEquals(styles.getHorizontal(style), Horizontals.CENTER);
                }
            }

            // 第七行
            assertTrue(iter.hasNext());
            assertTrue(iter.next().isBlank()); // 空行

            // 第八行
            assertTrue(iter.hasNext());
            row = iter.next();
            assertEquals(row.getFirstColumnIndex(), 0);
            assertEquals(row.getInt(0).intValue(), 1);
            assertEquals(row.getString(1), yzOrderEntity.jpCode);
            assertEquals(row.getString(3), yzOrderEntity.jpName);
            assertEquals(row.getInt(6).intValue(), yzOrderEntity.num);
            assertTrue(Math.abs(row.getDouble(7) - yzOrderEntity.price) <= 0.00001);
            assertTrue(Math.abs(row.getDouble(8) - yzOrderEntity.amount) <= 0.00001);
            assertTrue(Math.abs(row.getDouble(9) - yzOrderEntity.tax) <= 0.00001);
            assertTrue(Math.abs(row.getDouble(10) - yzOrderEntity.taxPrice) <= 0.00001);
            assertTrue(Math.abs(row.getDouble(11) - yzOrderEntity.taxAmount) <= 0.00001);
            assertEquals(row.getString(12), yzOrderEntity.remark);

            // 第九行
            assertTrue(iter.hasNext());
            assertTrue(iter.next().isBlank()); // 空行
            // 第十行
            assertTrue(iter.hasNext());
            assertTrue(iter.next().isBlank()); // 空行

            // 第十一行
            assertTrue(iter.hasNext());
            row = iter.next();
            assertEquals(row.getString(0), "合计");
            assertEquals(row.getInt(6).intValue(), yzSummary.nums);
            assertTrue(Math.abs(row.getDouble(7) - yzSummary.priceTotal) <= 0.00001);
            assertTrue(Math.abs(row.getDouble(8) - yzSummary.amountTotal) <= 0.00001);
            assertTrue(Math.abs(row.getDouble(9) - yzSummary.taxTotal) <= 0.00001);
            assertTrue(Math.abs(row.getDouble(10) - yzSummary.taxPriceTotal) <= 0.00001);
            assertTrue(Math.abs(row.getDouble(11) - yzSummary.taxAmountTotal) <= 0.00001);
        }
    }

    @Test public void testFillMap() throws IOException {
        final String fileName = "fill map.xlsx";
        Map<String, Object> master = YzEntity.mockMap(), summery = YzSummary.mockMap(), body = YzOrderEntity.randomMap().get(0);
        body.put("remark", "很长很长很长很长很长很长很长很长很长很长很长很长很长很长很长很长很长很长很长很长很长很长很长很长");
        new Workbook()
            .addSheet(new TemplateSheet(testResourceRoot().resolve("template2.xlsx"), "混合命名空间")
                    .setData(master)
                    // 单Map测试
                    .setData("YzEntity", body)
                    .setData("summary", summery)
            ).writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            Iterator<org.ttzero.excel.reader.Row> iter = reader.sheet(0).iterator();
            // 跳过前7行
            iter.next();iter.next(); iter.next();iter.next();iter.next();iter.next();iter.next();

            org.ttzero.excel.reader.Row row = iter.next();
            assertEquals(row.getString(12), "很长很长很长很长很长很长很长很长很长很长很长很长很长很长很长很长很长很长很长很长很长很长很长很长");
        }
    }

    @Test public void testFillListObject() throws IOException {
        final String fileName = "fill list object.xlsx";
        List<YzOrderEntity> expectList = YzOrderEntity.randomData();
        new Workbook()
            .addSheet(new TemplateSheet(testResourceRoot().resolve("template2.xlsx"), "混合命名空间")
                    .setData(YzEntity.mock())
                    .setData("YzEntity", expectList)
                    .setData("summary", YzSummary.mock())
            ).writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            assertListObject(reader.sheet(0).asFullSheet(), expectList);
        }
    }

    @Test public void testFillListMap() throws IOException {
        final String fileName = "fill list map.xlsx";
        List<Map<String, Object>> expectList = YzOrderEntity.randomMap();
        new Workbook()
            .addSheet(new TemplateSheet(testResourceRoot().resolve("template2.xlsx"), "混合命名空间")
                .setData(YzEntity.mockMap())
                .setData("YzEntity", expectList)
                .setData("summary", YzSummary.mockMap())
            ).writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            assertListMap(reader.sheet(0).asFullSheet(), expectList);
        }
    }

    @Test public void testFillSupplierListObject() throws IOException {
        final String fileName = "fill supplier list object.xlsx";
        List<YzOrderEntity> expectList = new ArrayList<>();
        new Workbook()
            .addSheet(new TemplateSheet(testResourceRoot().resolve("template2.xlsx"), "混合命名空间")
                .setData(YzEntity.mock())
                .setData("YzEntity", (i, o) -> {
                    List<YzOrderEntity> sub = null;
                    // 拉取100条数据
                    if (i < 100) {
                        YzOrderEntity lastOne = (YzOrderEntity) o;
                        sub = YzOrderEntity.randomData(lastOne != null ? lastOne.xh : 0);
                        expectList.addAll(sub);
                    }
                    return sub;
                })
                .setData("summary", YzSummary.mock())
            ).writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            assertListObject(reader.sheet(0).asFullSheet(), expectList);
        }
    }

    @Test public void testFillSupplierListMap() throws IOException {
        final String fileName = "fill supplier list map.xlsx";
        List<Map<String, Object>> expectList = new ArrayList<>();
        new Workbook()
            .addSheet(new TemplateSheet(testResourceRoot().resolve("template2.xlsx"), "混合命名空间")
                .setData(YzEntity.mockMap())
                .setData("YzEntity", (i, o) -> {
                    List<Map<String, Object>> sub = null;
                    // 拉取100条数据
                    if (i < 100) {
                        @SuppressWarnings("unchecked")
                        Map<String, Object> lastOne = (Map<String, Object>) o;
                        sub = YzOrderEntity.randomMap(lastOne != null ? Integer.parseInt(lastOne.get("xh").toString()) : 0);
                        expectList.addAll(sub);
                    }
                    return sub;
                })
                .setData("summary", YzSummary.mockMap())
            ).writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            assertListMap(reader.sheet(0).asFullSheet(), expectList);
        }
    }

    @Test public void testInnerFormula() throws IOException {
        final String fileName = "内置函数测试.xlsx";
        List<Map<String, Object>> list = new ArrayList<>();
        Map<String, Object> row1 = new HashMap<>();
        row1.put("name", "张三");
        row1.put("age", 6);
        row1.put("sex", "男");
        row1.put("pic", "https://gw.alicdn.com/bao/uploaded/i3/1081542738/O1CN01ZBcPlR1W63BQXG5yO_!!0-item_pic.jpg_300x300q90.jpg");
        row1.put("jumpUrl", "https://jianli.com/zhangsan");
        list.add(row1);

        Map<String, Object> row2 = new HashMap<>();
        row2.put("name", "李四");
        row2.put("age", 8);
        row2.put("sex", "女");
        row2.put("pic", "https://gw.alicdn.com/bao/uploaded/i3/2200754440203/O1CN01k8sRgC1DN1GGtuNT9_!!0-item_pic.jpg_300x300q90.jpg");
        row2.put("jumpUrl", "https://jianli.com/lisi");
        list.add(row2);

        new Workbook()
            // 模板工作表
            .addSheet(new TemplateSheet(testResourceRoot().resolve("template2.xlsx"), "内置函数")
                .setData(list)
                // 替换模板中"@list:sex"值为性别序列
                .setData("@list:sex", Arrays.asList("未知", "男", "女")))
            .writeTo(defaultTestPath.resolve(fileName));
    }

    @Ignore
    @Test public void test1kSheet() throws IOException {
        Workbook workbook = new Workbook();
        for (int i = 0; i < 1000; i++) {
            workbook.addSheet(new TemplateSheet(testResourceRoot().resolve("template2.xlsx"), "混合命名空间")
                .setData(YzEntity.mock())
                .setData("YzEntity", YzOrderEntity.mock(1000)));
        }
        final String fileName = "template 1k sheets.xlsx";
        workbook.writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            assertEquals(reader.all().length, 1000);
            for (int i = 0; i < 1000; i++) {
                org.ttzero.excel.reader.Sheet sheet = reader.sheet(i);
                assertEquals(sheet.rows().count(), 1010L);
            }
        }
    }

    @Ignore
    @Test public void test1kSheet2() throws IOException {
        final String fileName = "template 1k sheets.xlsx";
        AtomicInteger counter = new AtomicInteger(0);
        new Workbook().setWorkbookWriter(
            new SupplierXMLWorkbookWriter(() -> counter.incrementAndGet() <= 1000 ?
                new TemplateSheet(testResourceRoot().resolve("template2.xlsx"), "混合命名空间")
                    .setData(YzEntity.mock())
                    .setData("YzEntity", YzOrderEntity.mock(1000)) : null)
        ).writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            assertEquals(reader.all().length, 1000);
            for (int i = 0; i < 1000; i++) {
                org.ttzero.excel.reader.Sheet sheet = reader.sheet(i);
                assertEquals(sheet.rows().count(), 1010L);
            }
        }
    }

    @Test public void testDefaultFormatOnDateCell() throws IOException {
        Map<String, Object> data = new HashMap<>();
        data.put("channel", new Timestamp(System.currentTimeMillis()));
        new Workbook()
            .addSheet(new TemplateSheet(testResourceRoot().resolve("template2.xlsx")).setData(data))
            .writeTo(defaultTestPath.resolve("defaultFormatOnDateCell.xlsx"));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("defaultFormatOnDateCell.xlsx"))) {
            Row row = reader.sheet(0).header(1).iterator().next();
            int styleIndex = row.getCellStyle(0);
            NumFmt numFmt = reader.getStyles().getNumFmt(styleIndex);
            assertEquals(NumFmt.DATETIME_FORMAT, numFmt);
            Map<String, Object> map = row.toMap();
            assertEquals(map.get("渠道").getClass(), Timestamp.class);
            assertEquals(((Timestamp) map.get("渠道")).getTime() / 1000, ((Timestamp)data.get("channel")).getTime() / 1000);
        }
    }

    public static class SupplierXMLWorkbookWriter extends XMLWorkbookWriter {
        private final Supplier<Sheet> sheetSupplier;

        public SupplierXMLWorkbookWriter(Supplier<Sheet> sheetSupplier) {
            this.sheetSupplier = sheetSupplier;
        }

        @Override
        protected Path createTemp() throws IOException, ExcelWriteException {
            Workbook workbook = getWorkbook();
            Path root = null;
            try {
                root = FileUtil.mktmp(Const.EEC_PREFIX);
                Path xl = Files.createDirectory(root.resolve("xl"));

                ICellValueAndStyle cvas = new XMLCellValueAndStyle();
                // 在循环中创建sheet，创建好后就后就输出并回收
                for (int i = 1; i < 50000; i++) { // 为了安全最多限制5万个sheet
                    // 从supplier中获取sheet
                    Sheet sheet = sheetSupplier.get();
                    if (sheet == null) break;
                    sheet.setId(i);
                    if (StringUtil.isEmpty(sheet.getName())) {
                        sheet.setName("Sheet" + i);
                    }
                    sheet.setWorkbook(workbook);
                    sheet.setCellValueAndStyle(cvas);
                    sheet.setSheetWriter(getWorksheetWriter(sheet));
                    sheet.writeTo(xl);
                    sheet.close();

                    // 放入空的Sheet用于占位
                    workbook.addSheet(new ListSheet<>(sheet.getName()).setId(sheet.getId()));
                }

                writeGlobalAttribute(xl);
                Path zipFile = ZipUtil.zipExcludeRoot(root, workbook.getCompressionLevel(), root);
                FileUtil.rm_rf(root.toFile(), true);
                return zipFile;
            } catch (Exception e) {
                if (root != null) FileUtil.rm_rf(root);
                workbook.getSharedStrings().close();
                throw e;
            }
        }
    }

    static void assertListObject(FullSheet sheet, List<YzOrderEntity> expectList) {
        Iterator<org.ttzero.excel.reader.Row> iter = sheet.header(6, 7).iterator();

        List<Dimension> mergeCells = sheet.getMergeCells();
        assertEquals(mergeCells.size(), 26 + expectList.size() * 3);
        Map<Long, Dimension> mergeCellMap = new HashMap<>(mergeCells.size());
        for (Dimension dim : mergeCells) {
            mergeCellMap.put(TemplateSheet.dimensionKey(dim.firstRow - 1, dim.firstColumn - 1), dim);
        }
        org.ttzero.excel.reader.Row row;
        for (YzOrderEntity expect : expectList) {
            assertTrue(iter.hasNext());
            row = iter.next();
            assertEquals(row.getFirstColumnIndex(), 0);
            assertEquals(row.getInt(0).intValue(), expect.xh);
            assertEquals(row.getString(1), expect.jpCode);
            assertEquals(row.getString(3), expect.jpName);
            assertEquals(row.getInt(6).intValue(), expect.num);
            assertTrue(Math.abs(row.getDouble(7) - expect.price) <= 0.00001);
            assertTrue(Math.abs(row.getDouble(8) - expect.amount) <= 0.00001);
            assertTrue(Math.abs(row.getDouble(9) - expect.tax) <= 0.00001);
            assertTrue(Math.abs(row.getDouble(10) - expect.taxPrice) <= 0.00001);
            assertTrue(Math.abs(row.getDouble(11) - expect.taxAmount) <= 0.00001);
            assertEquals(row.getString(12), expect.remark);

            // 判断是否带合并
            Dimension mergeCell = mergeCellMap.get(TemplateSheet.dimensionKey(row.getRowNum() - 1, 1));
            assertNotNull(mergeCell);
            assertEquals(mergeCell.width, 2);
            mergeCell = mergeCellMap.get(TemplateSheet.dimensionKey(row.getRowNum() - 1, 3));
            assertNotNull(mergeCell);
            assertEquals(mergeCell.width, 3);
            mergeCell = mergeCellMap.get(TemplateSheet.dimensionKey(row.getRowNum() - 1, 12));
            assertNotNull(mergeCell);
            assertEquals(mergeCell.width, 5);
        }
        // 跳过2行
        assertTrue(iter.next().isBlank());
        assertTrue(iter.next().isBlank());

        // 合计行
        row = iter.next();
        assertEquals(row.getString(0), "合计");
        Dimension mergeCell = mergeCellMap.get(TemplateSheet.dimensionKey(row.getRowNum() - 1, 0));
        assertNotNull(mergeCell);
        assertEquals(mergeCell.width, 6);
        mergeCell = mergeCellMap.get(TemplateSheet.dimensionKey(row.getRowNum() - 1, 12));
        assertNotNull(mergeCell);
        assertEquals(mergeCell.width, 5);
    }

    static void assertListMap(FullSheet sheet, List<Map<String, Object>> expectList) {
        Iterator<org.ttzero.excel.reader.Row> iter = sheet.header(6, 7).iterator();

        List<Dimension> mergeCells = sheet.getMergeCells();
        assertEquals(mergeCells.size(), 26 + expectList.size() * 3);
        Map<Long, Dimension> mergeCellMap = new HashMap<>(mergeCells.size());
        for (Dimension dim : mergeCells) {
            mergeCellMap.put(TemplateSheet.dimensionKey(dim.firstRow - 1, dim.firstColumn - 1), dim);
        }
        org.ttzero.excel.reader.Row row;
        for (Map<String, Object> expect : expectList) {
            assertTrue(iter.hasNext());
            row = iter.next();
            assertEquals(row.getFirstColumnIndex(), 0);
            assertEquals(row.getInt(0), expect.get("xh"));
            assertEquals(row.getString(1), expect.get("jpCode"));
            assertEquals(row.getString(3), expect.get("jpName"));
            assertEquals(row.getInt(6), expect.get("num"));
            assertTrue(Math.abs(row.getDouble(7) - (double) expect.get("price")) <= 0.00001);
            assertTrue(Math.abs(row.getDouble(8) - (double) expect.get("amount")) <= 0.00001);
            assertTrue(Math.abs(row.getDouble(9) - (double) expect.get("tax")) <= 0.00001);
            assertTrue(Math.abs(row.getDouble(10) - (double) expect.get("taxPrice")) <= 0.00001);
            assertTrue(Math.abs(row.getDouble(11) - (double) expect.get("taxAmount")) <= 0.00001);
            assertEquals(row.getString(12), expect.get("remark"));

            // 判断是否带合并
            Dimension mergeCell = mergeCellMap.get(TemplateSheet.dimensionKey(row.getRowNum() - 1, 1));
            assertNotNull(mergeCell);
            assertEquals(mergeCell.width, 2);
            mergeCell = mergeCellMap.get(TemplateSheet.dimensionKey(row.getRowNum() - 1, 3));
            assertNotNull(mergeCell);
            assertEquals(mergeCell.width, 3);
            mergeCell = mergeCellMap.get(TemplateSheet.dimensionKey(row.getRowNum() - 1, 12));
            assertNotNull(mergeCell);
            assertEquals(mergeCell.width, 5);
        }
        // 跳过2行
        assertTrue(iter.next().isBlank());
        assertTrue(iter.next().isBlank());

        // 合计行
        row = iter.next();
        assertEquals(row.getString(0), "合计");
        Dimension mergeCell = mergeCellMap.get(TemplateSheet.dimensionKey(row.getRowNum() - 1, 0));
        assertNotNull(mergeCell);
        assertEquals(mergeCell.width, 6);
        mergeCell = mergeCellMap.get(TemplateSheet.dimensionKey(row.getRowNum() - 1, 12));
        assertNotNull(mergeCell);
        assertEquals(mergeCell.width, 5);
    }

    public static class YzEntity {
        private String gysName;
        private String gsName;
        private String jsName;
        private String orderNo;
        private String orderStatus;
        private Date cgDate;

        public static YzEntity mock() {
            YzEntity e = new YzEntity();
            e.gysName =" 供应商";
            e.gsName = "ABC公司";
            e.jsName = "亚瑟";
            e.cgDate = new Date();
            e.orderNo = "JD-0001";
            e.orderStatus = "OK";
            return e;
        }

        public static Map<String, Object> mockMap() {
            Map<String, Object> e = new HashMap<>();
            e.put("gysName", "供应商A");
            e.put("gsName", "ABC公司");
            e.put("jsName", "亚瑟");
            e.put("cgDate", new Date());
            e.put("orderNo", "JD-0001");
            e.put("orderStatus", "OK");
            return e;
        }
    }

    public static class YzSummary {
        private int nums;
        private double priceTotal;
        private double amountTotal;
        private double taxTotal;
        private double taxPriceTotal;
        private double taxAmountTotal;

        public static YzSummary mock() {
            YzSummary e = new YzSummary();
            e.nums = 10;
            e.priceTotal = 10;
            e.amountTotal = 10;
            e.taxTotal = 10;
            e.taxPriceTotal = 10;
            e.taxAmountTotal = 10;
            return e;
        }

        public static Map<String, Object> mockMap() {
            Map<String, Object> e = new HashMap<>();
            e.put("nums", 10);
            e.put("priceTotal", 10);
            e.put("amountTotal", 10);
            e.put("taxTotal", 10);
            e.put("taxPriceTotal", 10);
            e.put("taxAmountTotal", 10);
            return e;
        }
    }

    public static class YzOrderEntity {
        private int xh;
        private String jpCode;
        private String jpName;
        private int num;
        private double price;
        private double amount;
        private double tax;
        private double taxPrice;
        private double taxAmount;
        private String remark;

        private static YzOrderEntity mock() {
            YzOrderEntity e = new YzOrderEntity();
            e.xh = 1;
            e.jpCode = "code1";
            e.jpName = "name1";
            e.num = e.xh;
            e.price = 3.5D * e.xh;
            e.amount = e.price * e.num;
            e.tax = 0.006;
            e.taxPrice = e.price * (e.tax + 1);
            e.taxAmount = e.amount * e.tax;
            e.remark = "备注";
            return e;
        }

        public static List<YzOrderEntity> mock(int len) {
            List<YzOrderEntity> list = new ArrayList<>(len);
            for (int i = 0; i < len; i++) {
                YzOrderEntity e = new YzOrderEntity();
                e.xh = i + 1;
                e.jpCode = "code" + e.xh;
                e.jpName = "name" + e.xh;
                e.num = e.xh;
                e.price = 3.5D * e.xh;
                e.amount = e.price * e.num;
                e.tax = 0.006;
                e.taxPrice = e.price * (e.tax + 1);
                e.taxAmount = e.amount * e.tax;
                e.remark = "备注" + e.xh;
                list.add(e);
            }
            return list;
        }

        public static List<YzOrderEntity> randomData() {
            return randomData(0);
        }

        public static List<YzOrderEntity> randomData(int startIndex) {
            List<YzOrderEntity> list = new ArrayList<>(10);
            for (int i = Math.max(startIndex, 0), len = i + 10; i < len; i++) {
                YzOrderEntity e = new YzOrderEntity();
                e.xh = i + 1;
                e.jpCode = "code" + e.xh;
                e.jpName = "name" + e.xh;
                e.num = e.xh;
                e.price = 3.5D * e.xh;
                e.amount = e.price * e.num;
                e.tax = 0.006;
                e.taxPrice = e.price * (e.tax + 1);
                e.taxAmount = e.amount * e.tax;
                e.remark = "备注" + e.xh;
                list.add(e);
            }
            return list;
        }

        public static List<Map<String, Object>> randomMap() {
            return randomMap(0);
        }

        public static List<Map<String, Object>> randomMap(int startIndex) {
            List<Map<String, Object>> list = new ArrayList<>(10);
            for (int i = Math.max(startIndex, 0), len = i + 10, j; i < len; i++) {
                Map<String, Object> map = new HashMap<>();
                map.put("xh", (j = i + 1));
                map.put("jpCode", "code" + j);
                map.put("jpName", "name" + j);
                map.put("num", j);
                map.put("price", 3.5D * j);
                map.put("amount", 3.5D * j * j);
                map.put("tax", 0.006);
                map.put("taxPrice", 3.5D * j * 1.006);
                map.put("taxAmount", 3.5D * j * 1.006 * j);
                map.put("remark", "备注" + j);
                list.add(map);
            }
            return list;
        }
    }
    public static class GameEntry {
        Integer channel;
        String game, account, vip;
        java.util.Date date;
        Boolean isAdult;

        public static List<GameEntry> random() {
            List<GameEntry> list = new ArrayList<>(10);
            String[] games = { "LOL", "WOW", "守望先锋", "怪物世界", "极品飞车" };
            for (int i = 0, v; i < 10; i++) {
                GameEntry e = new GameEntry();
                list.add(e);
                e.channel = random.nextInt(10);
                e.game = games[random.nextInt(games.length)];
                e.account = getRandomAssicString(10);
                e.date = new Timestamp(System.currentTimeMillis() - random.nextInt(1000000));
                e.isAdult = random.nextInt(10) <= 2;
                e.vip = (v = random.nextInt(100)) < 15 ? "v" + ((v & 3) + 1) : null;
            }
            return list;
        }

        @Override
        public boolean equals(Object o) {
            if (o instanceof GameEntry) {
                GameEntry e = (GameEntry) o;
                return Objects.equals(channel, e.channel)
                    && Objects.equals(isAdult, e.isAdult)
                    && Objects.equals(game, e.game)
                    && Objects.equals(account, e.account)
                    && Objects.equals(vip, e.vip)
                    && date.getTime() / 1000 == e.date.getTime() / 1000;
            }
            return false;
        }
    }
}

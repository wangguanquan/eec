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

import org.junit.Test;
import org.ttzero.excel.annotation.ExcelColumn;
import org.ttzero.excel.annotation.Hyperlink;
import org.ttzero.excel.entity.e7.XMLWorksheetWriter;
import org.ttzero.excel.entity.style.Border;
import org.ttzero.excel.entity.style.BorderStyle;
import org.ttzero.excel.entity.style.Fill;
import org.ttzero.excel.entity.style.Font;
import org.ttzero.excel.entity.style.Horizontals;
import org.ttzero.excel.entity.style.PatternType;
import org.ttzero.excel.entity.style.Styles;
import org.ttzero.excel.manager.Const;
import org.ttzero.excel.processor.Converter;
import org.ttzero.excel.processor.StyleProcessor;
import org.ttzero.excel.reader.Cell;
import org.ttzero.excel.reader.Dimension;
import org.ttzero.excel.reader.ExcelReader;
import org.ttzero.excel.reader.FullSheet;
import org.ttzero.excel.reader.Grid;
import org.ttzero.excel.reader.GridFactory;
import org.ttzero.excel.reader.HeaderRow;
import org.ttzero.excel.reader.Sheet;
import org.ttzero.excel.util.StringUtil;

import java.awt.Color;
import java.io.IOException;
import java.lang.annotation.Documented;
import java.lang.annotation.ElementType;
import java.lang.annotation.Inherited;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;
import java.lang.reflect.AccessibleObject;
import java.lang.reflect.Array;
import java.math.BigDecimal;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.HashMap;
import java.util.Objects;
import java.util.stream.Collectors;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;
import static org.ttzero.excel.entity.Sheet.int2Col;
import static org.ttzero.excel.util.StringUtil.isNotEmpty;

/**
 * @author guanquan.wang at 2023-04-04 22:38
 */
public class ListObjectSheetTest2 extends WorkbookTest {
    @Test public void testSpecifyRowWrite() throws IOException {
        List<ListObjectSheetTest.Item> list = ListObjectSheetTest.Item.randomTestData();
        new Workbook().setAutoSize(true)
            .addSheet(new ListSheet<>(list).setStartRowIndex(4))
            .writeTo(defaultTestPath.resolve("test specify row 4 ListSheet.xlsx"));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("test specify row 4 ListSheet.xlsx"))) {
            List<ListObjectSheetTest.Item> readList = reader.sheet(0).header(4).rows().map(row -> row.to(ListObjectSheetTest.Item.class)).collect(Collectors.toList());
            assertEquals(list.size(), readList.size());
            for (int i = 0, len = list.size(); i < len; i++)
                assertEquals(list.get(i), readList.get(i));
        }
    }

    @Test public void testSpecifyRowStayA1Write() throws IOException {
        List<ListObjectSheetTest.Item> list = ListObjectSheetTest.Item.randomTestData();
        new Workbook().setAutoSize(true)
            .addSheet(new ListSheet<>(list).setStartRowIndex(4, false))
            .writeTo(defaultTestPath.resolve("test specify row 4 stay A1 ListSheet.xlsx"));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("test specify row 4 stay A1 ListSheet.xlsx"))) {
            List<ListObjectSheetTest.Item> readList = reader.sheet(0).bind(ListObjectSheetTest.Item.class, 4).rows().map(row -> (ListObjectSheetTest.Item) row.get()).collect(Collectors.toList());
            assertEquals(list.size(), readList.size());
            for (int i = 0, len = list.size(); i < len; i++)
                assertEquals(list.get(i), readList.get(i));
        }
    }

    @Test public void testSpecifyRowAndColWrite() throws IOException {
        List<ListObjectSheetTest.Item> list = ListObjectSheetTest.Item.randomTestData(10);
        new Workbook().setAutoSize(true)
            .addSheet(new ListSheet<ListObjectSheetTest.Item>("Item"
                , new Column("id").setColIndex(3)
                , new Column("name").setColIndex(4))
                .setData(list)
                .setStartRowIndex(4)
            ).writeTo(defaultTestPath.resolve("test specify row and cel ListSheet.xlsx"));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("test specify row and cel ListSheet.xlsx"))) {
            List<ListObjectSheetTest.Item> readList = reader.sheet(0).bind(ListObjectSheetTest.Item.class, 4).rows().map(row -> (ListObjectSheetTest.Item) row.get()).collect(Collectors.toList());
            assertEquals(list.size(), readList.size());
            for (int i = 0, len = list.size(); i < len; i++)
                assertEquals(list.get(i), readList.get(i));
        }
    }

    @Test public void testSpecifyRowAndColStayA1Write() throws IOException {
        List<ListObjectSheetTest.Item> list = ListObjectSheetTest.Item.randomTestData(10);
        new Workbook().setAutoSize(true)
            .addSheet(new ListSheet<ListObjectSheetTest.Item>("Item"
                , new Column("id").setColIndex(3)
                , new Column("name").setColIndex(4))
                .setData(list)
                .setStartRowIndex(4, false)
            ).writeTo(defaultTestPath.resolve("test specify row and cel stay A1 ListSheet.xlsx"));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("test specify row and cel stay A1 ListSheet.xlsx"))) {
            List<ListObjectSheetTest.Item> readList = reader.sheet(0).bind(ListObjectSheetTest.Item.class, 4).rows().map(row -> (ListObjectSheetTest.Item) row.get()).collect(Collectors.toList());
            assertEquals(list.size(), readList.size());
            for (int i = 0, len = list.size(); i < len; i++)
                assertEquals(list.get(i), readList.get(i));
        }
    }

    @Test public void testSpecifyRowIgnoreHeaderWrite() throws IOException {
        List<ListObjectSheetTest.Item> list = ListObjectSheetTest.Item.randomTestData();
        new Workbook().setAutoSize(true)
            .addSheet(new ListSheet<>(list).setStartRowIndex(4).ignoreHeader())
            .writeTo(defaultTestPath.resolve("test specify row 4 ignore header ListSheet.xlsx"));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("test specify row 4 ignore header ListSheet.xlsx"))) {
            List<ListObjectSheetTest.Item> readList = reader.sheet(0)
                .header(3)
                .bind(ListObjectSheetTest.Item.class, new HeaderRow().with(createHeaderRow()))
                .rows()
                .map(row -> (ListObjectSheetTest.Item) row.get())
                .collect(Collectors.toList());
            assertEquals(list.size(), readList.size());
            for (int i = 0, len = list.size(); i < len; i++)
                assertEquals(list.get(i), readList.get(i));
        }
    }

    @Test public void testSpecifyRowStayA1IgnoreHeaderWrite() throws IOException {
        List<ListObjectSheetTest.Item> list = ListObjectSheetTest.Item.randomTestData();
        new Workbook().setAutoSize(true)
            .addSheet(new ListSheet<>(list).setStartRowIndex(4, false).ignoreHeader())
            .writeTo(defaultTestPath.resolve("test specify row 4 stay A1 ignore header ListSheet.xlsx"));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("test specify row 4 stay A1 ignore header ListSheet.xlsx"))) {
            List<ListObjectSheetTest.Item> readList = reader.sheet(0).rows().map(row -> {
                ListObjectSheetTest.Item e = new ListObjectSheetTest.Item();
                e.setId(row.getInt(0));
                e.setName(row.getString(1));
                return e;
            }).collect(Collectors.toList());
            assertEquals(list.size(), readList.size());
            for (int i = 0, len = list.size(); i < len; i++)
                assertEquals(list.get(i), readList.get(i));
        }
    }

    @Test public void testSpecifyRowAndColIgnoreHeaderWrite() throws IOException {
        List<ListObjectSheetTest.Item> list = ListObjectSheetTest.Item.randomTestData(10);
        new Workbook().setAutoSize(true)
            .addSheet(new ListSheet<ListObjectSheetTest.Item>("Item"
                , new Column("id").setColIndex(3)
                , new Column("name").setColIndex(4))
                .setData(list)
                .setStartRowIndex(4)
                .ignoreHeader()
            ).writeTo(defaultTestPath.resolve("test specify row and cel ignore header ListSheet.xlsx"));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("test specify row and cel ignore header ListSheet.xlsx"))) {
            List<ListObjectSheetTest.Item> readList = reader.sheet(0).rows().map(row -> {
                ListObjectSheetTest.Item e = new ListObjectSheetTest.Item();
                e.setId(row.getInt(3));
                e.setName(row.getString(4));
                return e;
            }).collect(Collectors.toList());
            assertEquals(list.size(), readList.size());
            for (int i = 0, len = list.size(); i < len; i++)
                assertEquals(list.get(i), readList.get(i));
        }
    }

    @Test public void testSpecifyRowAndColStayA1IgnoreHeaderWrite() throws IOException {
        List<ListObjectSheetTest.Item> list = ListObjectSheetTest.Item.randomTestData(10);
        new Workbook().setAutoSize(true)
            .addSheet(new ListSheet<ListObjectSheetTest.Item>("Item"
                , new Column("id").setColIndex(3)
                , new Column("name").setColIndex(4))
                .setData(list)
                .setStartRowIndex(4, false)
                .ignoreHeader()
            ).writeTo(defaultTestPath.resolve("test specify row and cel stay A1 ignore header ListSheet.xlsx"));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("test specify row and cel stay A1 ignore header ListSheet.xlsx"))) {
            List<ListObjectSheetTest.Item> readList = reader.sheet(0).rows().map(row -> {
                ListObjectSheetTest.Item e = new ListObjectSheetTest.Item();
                e.setId(row.getInt(3));
                e.setName(row.getString(4));
                return e;
            }).collect(Collectors.toList());
            assertEquals(list.size(), readList.size());
            for (int i = 0, len = list.size(); i < len; i++)
                assertEquals(list.get(i), readList.get(i));
        }
    }

    @Test public void testCustomerRowHeight() throws IOException {
        List<Template> list = new ArrayList<>();
        list.add(Template.of("备注说明\r\n第二行\r\n第三行\r\n第四行", "岗位名称", "岁位"));
        list.add(Template.of("字段名称", "*岗位名称", "岗位描述"));
        list.add(Template.of("示例", "生产统计员", "按照产品规格、价格、工序、员工、车间等不同对象和要求进行统计数据资料分析"));

        new Workbook().addSheet(
            new ListSheet<>(list).setStyleProcessor(new TemplateStyleProcessor())
                .setRowHeight(62.25D)
                .cancelZebraLine().ignoreHeader().putExtProp(Const.ExtendPropertyKey.MERGE_CELLS, Collections.singletonList(Dimension.of("A1:B1")))
        ).writeTo(defaultTestPath.resolve("Customer row height.xlsx"));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("Customer row height.xlsx"))) {
            Iterator<org.ttzero.excel.reader.Row> iter = reader.sheet(0).rows().iterator();
            assertTrue(iter.hasNext());
            org.ttzero.excel.reader.Row row0 = iter.next();
            assertEquals(list.get(0), Template.of(row0.getString(0), row0.getString(1), row0.getString(2)));
            Styles styles = row0.getStyles();
            int styleIndex = row0.getCellStyle(0);
            Fill fill0 = styles.getFill(styleIndex), fill1 = styles.getFill(row0.getCellStyle(1)), fill2 = styles.getFill(row0.getCellStyle(2));
            assertTrue(fill0 != null && fill0.getPatternType() == PatternType.solid && fill0.getFgColor().equals(new Color(188, 219, 162)));
            assertTrue(fill1 == null || fill1.getPatternType() == PatternType.none);
            assertTrue(fill2 == null || fill2.getPatternType() == PatternType.none);

            assertTrue(iter.hasNext());
            org.ttzero.excel.reader.Row row1 = iter.next();
            assertEquals(list.get(1), Template.of(row1.getString(0), row1.getString(1), row1.getString(2)));
            org.ttzero.excel.entity.style.Font font0 = styles.getFont(row1.getCellStyle(0)), font1 = styles.getFont(row1.getCellStyle(1)), font2 = styles.getFont(row1.getCellStyle(2));
            assertTrue(font0.isBold());
            assertTrue(font1.isBold());
            assertTrue(font2.isBold());
            assertEquals(styles.getHorizontal(row1.getCellStyle(0)), Horizontals.LEFT);
            assertEquals(styles.getHorizontal(row1.getCellStyle(1)), Horizontals.CENTER);
            assertEquals(styles.getHorizontal(row1.getCellStyle(2)), Horizontals.CENTER);

            assertTrue(iter.hasNext());
            org.ttzero.excel.reader.Row row2 = iter.next();
            assertEquals(list.get(2), Template.of(row2.getString(0), row2.getString(1), row2.getString(2)));
            assertEquals(styles.getHorizontal(row2.getCellStyle(0)), Horizontals.LEFT);
            assertEquals(styles.getHorizontal(row2.getCellStyle(1)), Horizontals.CENTER);
            assertEquals(styles.getHorizontal(row2.getCellStyle(2)), Horizontals.LEFT);
        }
    }

    @Test public void testTileWriter() throws IOException {
        String fileName = "Dynamic title.xlsx";
        List<TileEntity> data = TileEntity.randomTestData();
        new Workbook().cancelZebraLine().addSheet(new ListSheet<>(data).setSheetWriter(new TileXMLWorksheetWriter(3, LocalDate.now().toString()))).writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            Iterator<org.ttzero.excel.reader.Row> iter = reader.sheet(0).rows().iterator();
            assertTrue(iter.hasNext());
            assertEquals((LocalDate.now() +  " 拣货单"), iter.next().getString(0));

            assertTrue(iter.hasNext());
            assertEquals("差异 | 序号 | 商品 | 数量 | 差异 | 序号 | 商品 | 数量 | 差异 | 序号 | 商品 | 数量", iter.next().toString());

            // TODO assert row data
        }
    }

    @Test public void testEmptySheetSubClassSpecified() throws IOException {
        String fileName = "sub-class specified types.xlsx";
        List<ListObjectSheetTest.Item> expectList = new ArrayList<>();
        new Workbook().setAutoSize(true)
            .addSheet(new ListSheet<ListObjectSheetTest.Item>() {
                int i = 0;
                @Override
                protected List<ListObjectSheetTest.Item> more() {
                    List<ListObjectSheetTest.Item> list = i++ < 1 ? ListObjectSheetTest.Item.randomTestData(10) : null;
                    if (list != null) expectList.addAll(list);
                    return list;
                }
            })
            .writeTo(defaultTestPath.resolve(fileName));

        // Check header row
        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            List<ListObjectSheetTest.Item> list = reader.sheet(0).dataRows().map(row -> row.to(ListObjectSheetTest.Item.class)).collect(Collectors.toList());
            assertEquals(list.size(), expectList.size());
            for (int i = 0, len = expectList.size(); i < len; i++) {
               ListObjectSheetTest.Item expect = expectList.get(i), e = list.get(i);
               assertEquals(expect, e);
            }
        }
    }

    @Test public void testSpecifyActualClass() throws IOException {
        String fileName = "specify unrelated class.xlsx";
        new Workbook()
            .addSheet(new ListSheet<>().setClass(SubModel.class))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            Iterator<org.ttzero.excel.reader.Row> iter = reader.sheet(0).iterator();
            assertTrue(iter.hasNext());
            org.ttzero.excel.reader.Row row = iter.next();
            assertEquals("name", row.getString(0));
            assertEquals("status", row.getString(1));
        }
    }

    @Test public void testSpecifyConvertClass() throws IOException {
        List<SpecifyConvertModel> expectList = SpecifyConvertModel.randomTestData(20);
        String fileName = "specify converter test.xlsx";
        new Workbook()
            .addSheet(new ListSheet<>(expectList))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            List<SpecifyConvertModel> readList = reader.sheet(0).header(1).rows().map(row -> row.to(SpecifyConvertModel.class)).collect(Collectors.toList());
            assertEquals(expectList.size(), readList.size());
            for (int i = 0, len = expectList.size(); i < len; i++)
                assertEquals(expectList.get(i), readList.get(i));
        }
    }

    @Test public void testAutoSize() throws IOException {
        String fileName = "test auto size.xlsx";
        List<ListObjectSheetTest.Student> expectList = ListObjectSheetTest.Student.randomTestData();
        new Workbook()
            .setAutoSize(true)
            .addSheet(new ListSheet<>(expectList
                , new Column("学号", "id").setStyleProcessor((o, style, sst)
                    -> (((int) o & 1) == 1 ? sst.modifyFont(style, new Font("Algerian", 24)) : ((int) o) < 10 ? sst.modifyFont(style, new Font("Algerian", 56)) : style))
                , new Column("姓名", "name").setStyleProcessor((o, style, sst) -> {
                    int len = ((String) o).length();
                    if (len < 5) {
                        style = sst.modifyFont(style, new Font("Trebuchet MS", 72));
                    } else if (len > 15) {
                        style = sst.modifyFont(style, new Font("宋体", 5));
                    } else if (len > 10) {
                        style = sst.modifyFont(style, new Font("Bauhaus 93", 18));
                    }
                    return style;
                })
            ))
            .writeTo(defaultTestPath.resolve(fileName));
    }

    @Test public void testCustomStyle() throws IOException {
        String fileName = "test custom style.xlsx";
        List<ListObjectSheetTest.Student> expectList = ListObjectSheetTest.Student.randomTestData();
        new Workbook()
            .addSheet(new ListSheet<>(expectList
                , new Column("id").setFont(new Font("微软雅黑", 16)).setHorizontal(Horizontals.CENTER)
                , new Column("name").setFont(new Font("华文行楷", 23)).setBorder(new Border()).autoSize()
            ))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            Sheet sheet = reader.sheet(0);
            List<ListObjectSheetTest.Student> list = sheet.forceImport().dataRows().map(row -> row.to(ListObjectSheetTest.Student.class)).collect(Collectors.toList());
            assertEquals(list.size(), expectList.size());
            for (int i = 0; i < expectList.size(); i++) {
                ListObjectSheetTest.Student e = expectList.get(i), o = list.get(i);
                assertEquals(e.getName(), o.getName());
                assertEquals(e.getId(), o.getId());
            }

            for (Iterator<org.ttzero.excel.reader.Row> iter = sheet.reset().dataRows().iterator(); iter.hasNext(); ) {
                org.ttzero.excel.reader.Row row = iter.next();
                Styles styles = row.getStyles();
                // 第一列样式
                {
                    int style = row.getCellStyle(0);
                    Font font = styles.getFont(style);
                    assertEquals("微软雅黑", font.getName());
                    assertEquals(font.getSize(), 16);
                    int horizontal = styles.getHorizontal(style);
                    assertEquals(horizontal, Horizontals.CENTER);
                }
                // 第二列样式
                {
                    int style = row.getCellStyle(1);
                    Font font = styles.getFont(style);
                    assertEquals("华文行楷", font.getName());
                    assertEquals(font.getSize(), 23);
                    int horizontal = styles.getHorizontal(style);
                    assertEquals(horizontal, Horizontals.LEFT);
                    Border border = styles.getBorder(style);
                    assertTrue(border == null || border.getBorderTop().getStyle() == BorderStyle.NONE);
                }
            }
        }
    }

    @Test public void testSpecifyRowLimit() throws IOException {
        final String fileName = "specify row limit.xlsx";
        List<ListObjectSheetTest.Student> expectList = ListObjectSheetTest.Student.randomTestData(1000);
        new Workbook().addSheet(new ListSheet<>(expectList).setSheetWriter(new XMLWorksheetWriter() {
            @Override
            public int getRowLimit() {
                return 150;
            }
        })).writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            assertEquals(reader.getSheetCount(), 7);
            List<ListObjectSheetTest.Student> readList = reader.sheets().flatMap(Sheet::dataRows).map(row -> row.to(ListObjectSheetTest.Student.class)).collect(Collectors.toList());
            assertEquals(expectList.size(), readList.size());
            for (int i = 0, len = expectList.size(); i < len; i++) {
                ListObjectSheetTest.Student expect = expectList.get(i), o = readList.get(i);
                assertEquals(expect.getName(), o.getName());
                assertEquals(expect.getScore(), o.getScore());
            }
        }
    }

    @Test public void testAutoFilter() throws IOException {
        String fileName = "test auto-filter.xlsx";
        List<ListObjectSheetTest.Student> expectList = ListObjectSheetTest.Student.randomTestData();
        new Workbook()
            .addSheet(new ListSheet<>(expectList
                , new Column("学号", "id")
                , new Column("姓名", "name")
                , new Column("成绩", "score", n -> (int) n < 60 ? "不合格" : n)
            ).putExtProp(Const.ExtendPropertyKey.AUTO_FILTER, Dimension.of("A1:C1")))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            org.ttzero.excel.reader.FullSheet sheet = (FullSheet) reader.sheet(0).asFullSheet().header(1);
            org.ttzero.excel.reader.HeaderRow header = (HeaderRow) sheet.getHeader();
            assertEquals("学号", header.get(0));
            assertEquals("姓名", header.get(1));
            assertEquals("成绩", header.get(2));

            assertEquals(Dimension.of("A1:C1"), sheet.getFilter());


            Iterator<org.ttzero.excel.reader.Row> iter = sheet.iterator();
            for (ListObjectSheetTest.Student expect : expectList) {
                assertTrue(iter.hasNext());
                Map<String, Object> e = iter.next().toMap();
                assertEquals(expect.getId(), Integer.parseInt(e.get("学号").toString()));
                assertEquals(expect.getName(), e.get("姓名").toString());
                if (expect.getScore() < 60) {
                    assertEquals("不合格", e.get("成绩"));
                } else {
                    assertEquals(expect.getScore(), Integer.parseInt(e.get("成绩").toString()));
                }
            }
        }
    }

    @Test public void testAllNullObject() throws IOException {
        final String fileName = "all null object.xlsx";
        List<ListObjectSheetTest.Item> expectList = new ArrayList<>();
        expectList.add(null);
        expectList.add(null);
        expectList.add(null);
        expectList.add(null);
        expectList.add(null);
        new Workbook().addSheet(new ListSheet<>(expectList)).writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            assertEquals(reader.sheet(0).rows().count(), 0);
        }
    }

    public static class TemplateStyleProcessor implements StyleProcessor<Template> {
        String k;
        int c = 0;
        @Override
        public int build(Template o, int style, Styles sst) {
            if (!o.v1.equals(k)) {
                k = o.v1;
                c = 0;
            }
            if (o.v1.startsWith("备注说明")) {
                if (c == 0)
                    style = sst.modifyFill(style, new Fill(PatternType.solid, new Color(188, 219, 162)));
            }
            else if (o.v1.equals("字段名称")) {
                Font font = sst.getFont(style);
                style = sst.modifyFont(style, font.clone().bold());
                if (c > 0)
                    style = sst.modifyHorizontal(style, Horizontals.CENTER);
            }
            else if (o.v1.equals("示例")) {
                if (c == 1)
                    style = sst.modifyHorizontal(style, Horizontals.CENTER);
            }
            c++;
            return style;
        }
    }

    @Test public void testDataSupplier() throws IOException {
        final String fileName = "list data supplier.xlsx";
        List<ListObjectSheetTest.Student> expectList = new ArrayList<>(100);
        new Workbook()
            .addSheet(new ListSheet<ListObjectSheetTest.Student>().setData((i, lastOne) -> {
                if (i >= 100) return null;
                List<ListObjectSheetTest.Student> sub = ListObjectSheetTest.Student.randomTestData();
                expectList.addAll(sub);
                return sub;
            }))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            List<ListObjectSheetTest.Student> list =  reader.sheet(0).dataRows().map(row -> row.to(ListObjectSheetTest.Student.class)).collect(Collectors.toList());
            assertEquals(expectList.size(), list.size());
            for (int i = 0, len = expectList.size(); i < len; i++) {
                ListObjectSheetTest.Student expect = expectList.get(i), e = list.get(i);
                expect.setId(0); // ID not exported
                assertEquals(expect, e);
            }
        }
    }

    @Test public void testTreeStyle() throws IOException {
        List<TreeNode> root = new ArrayList<>();
        TreeNode class1 = new TreeNode("一年级", (94 + 97) / 2.0D);
        root.add(class1);
        class1.children = (Arrays.asList(new TreeNode("张一", 94), new TreeNode("李一", 97)));
        TreeNode class2 = new TreeNode("二年级", (75 + 100 + 90) / 3.0D);
        root.add(class2);
        class2.children = (Arrays.asList(new TreeNode("张二", 75), new TreeNode("李二", 100), new TreeNode("王二", 90)));

        new Workbook().addSheet(new ListSheet<TreeNode>(root) {
            @Override
            protected EntryColumn createColumn(AccessibleObject ao) {
                EntryColumn column = super.createColumn(ao);
                if (column == null && ao.isAnnotationPresent(TreeLevel.class)) {
                    column = new EntryColumn();
                    column.setColIndex(99); // <- 设置一个不存在特殊列
                }
                return column;
            }

            @Override
            protected void mergeGlobalSetting(Class<?> clazz) {
                super.mergeGlobalSetting(clazz);
                if (clazz.isAnnotationPresent(TreeStyle.class)) {
                    putExtProp("tree_style", "1");
                }
            }

            @Override
            protected void calculateRealColIndex() {
                super.calculateRealColIndex();
                // 将上面设置的特殊列号改到尾列
                columns[columns.length - 1].getTail().colIndex = columns[columns.length - 2].getTail().colIndex + 1;
                columns[columns.length - 1].getTail().realColIndex = columns[columns.length - 2].getTail().realColIndex + 1;
            }

            // 将树结构降维，如果由level区分等级则不需要这一步
            @Override
            public void resetBlockData() {
                if (!eof && left() < rowBlock.capacity()) {
                    append();
                }
                // EOF
                int left = left();
                if (left == 0) return;
                List<TreeNode> nodes = new ArrayList<>(left);
                for (TreeNode e : data) {
                    nodes.add(e);
                    e.level = 0;
                    List<TreeNode> sub = e.children;
                    e.children = null;
                    for (TreeNode o : sub) {
                        nodes.add(o);
                        o.level = 1;
                        o.children = null;
                    }
                }
                this.data = nodes; // <- 替换原有数据
                this.start = 0;
                this.end += nodes.size() - left; // <- 重置尾下标

                super.resetBlockData();
            }
        }.setSheetWriter(new XMLWorksheetWriter() {
            boolean isTreeStyle;
            @Override
            protected void writeBefore() throws IOException {
                super.writeBefore();

                isTreeStyle = "1".equals(sheet.getExtPropValue("tree_style"));
            }

            protected int startRow(int rows, int columns, Double rowHeight, int level) throws IOException {
                // Row number
                int r = rows + startRow;

                bw.write("<row r=\"");
                bw.writeInt(r);
                // default data row height 16.5
                if (rowHeight != null && rowHeight >= 0D) {
                    bw.write("\" customHeight=\"1\" ht=\"");
                    bw.write(rowHeight);
                }
                if (this.columns.length > 0) {
                    bw.write("\" spans=\"");
                    bw.writeInt(this.columns[0].realColIndex);
                    bw.write(':');
                    bw.writeInt(this.columns[this.columns.length - 1].realColIndex);
                } else {
                    bw.write("\" spans=\"1:");
                    bw.writeInt(columns);
                }
                if (level > 0) {
                    bw.write("\" outlineLevel=\"");
                    bw.writeInt(level);
                }
                bw.write("\">");
                return r;
            }

            @Override
            protected int writeHeaderRow() throws IOException {
                // Write header
                int rowIndex = 0, subColumnSize = columns[0].subColumnSize(), defaultStyleIndex = sheet.defaultHeadStyleIndex();
                int realColumnLen = isTreeStyle ? columns.length - 1 : columns.length;
                Column[][] columnsArray = new Column[realColumnLen][];
                for (int i = 0; i < realColumnLen; i++) {
                    columnsArray[i] = columns[i].toArray();
                }
                // Merge cells if exists
                @SuppressWarnings("unchecked")
                List<Dimension> mergeCells = (List<Dimension>) sheet.getExtPropValue(Const.ExtendPropertyKey.MERGE_CELLS);
                Grid mergedGrid = mergeCells != null && !mergeCells.isEmpty() ? GridFactory.create(mergeCells) : null;
                Cell cell = new Cell();
                for (int i = subColumnSize - 1; i >= 0; i--) {
                    // Custom row height
                    double ht = getHeaderHeight(columnsArray, i);
                    if (ht < 0) ht = sheet.getHeaderRowHeight();
                    int row = startRow(rowIndex++, realColumnLen, ht);

                    String name;
                    for (int j = 0, c = 0; j < realColumnLen; j++) {
                        Column hc = columnsArray[j][i];
                        cell.setString(isNotEmpty(hc.getName()) ? hc.getName() : mergedGrid != null && mergedGrid.test(i + 1, hc.getRealColIndex()) && !isFirstMergedCell(mergeCells, i + 1, hc.getRealColIndex()) ? null : hc.key);
                        cell.xf = hc.getHeaderStyleIndex() == -1 ? defaultStyleIndex : hc.getHeaderStyleIndex();
                        writeString(cell, row, c++);
                    }

                    // Write header comments
                    for (int j = 0; j < realColumnLen; j++) {
                        Column hc = columnsArray[j][i];
                        if (hc.headerComment != null) {
                            if (comments == null) comments = sheet.createComments();
                            comments.addComment(new String(int2Col(hc.getRealColIndex())) + row, hc.headerComment);
                        }
                    }
                    bw.write("</row>");
                }
                return subColumnSize;
            }

            @Override
            protected void writeRow(Row row) throws IOException {
                Cell[] cells = row.getCells();
                int len = isTreeStyle ? cells.length - 1 : cells.length;
                int r = isTreeStyle ? startRow(row.getIndex(), len, row.getHeight(), cells[columns.length - 1].intVal) : startRow(row.getIndex(), len, row.getHeight());

                for (int i = row.fc; i < row.lc; i++) writeCell(cells[i], r, i);

                bw.write("</row>");
            }
        })).writeTo(defaultTestPath.resolve("tree style.xlsx"));
    }

    @TreeStyle
    public static class TreeNode {
        @ExcelColumn
        String name;
        @ExcelColumn
        double score; // <- root节点表示平均成绩
        @TreeLevel
        int level; // <- 层级
        public TreeNode() { }
        public TreeNode(String name, double score) {
            this.name = name;
            this.score = score;
        }

        List<TreeNode> children;
    }

    @Target({ ElementType.TYPE })
    @Retention(RetentionPolicy.RUNTIME)
    @Inherited
    @Documented
    public @interface TreeStyle { }

    @Target({ ElementType.FIELD, ElementType.METHOD })
    @Retention(RetentionPolicy.RUNTIME)
    @Inherited
    @Documented
    public @interface TreeLevel { }

    public static class TileEntity {
        @ExcelColumn("{date} 拣货单")
        @ExcelColumn(value = "差异", maxWidth = 8.6D)
        private String diff;
        @ExcelColumn("{date} 拣货单")
        @ExcelColumn(value = "序号", maxWidth = 6.8D)
        private Integer no;
        @ExcelColumn("{date} 拣货单")
        @ExcelColumn(value = "商品", maxWidth = 12.0D)
        private String product;
        @ExcelColumn("{date} 拣货单")
        @ExcelColumn(value = "数量", maxWidth = 6.8D)
        private Integer num;

        public static List<TileEntity> randomTestData() {
            int n = 23;
            List<TileEntity> list = new ArrayList<>(n);
            for (int i = 0; i < n; i++) {
                TileEntity e = new TileEntity();
                e.no = i + 1;
                e.product = getRandomString(10);
                e.num = random.nextInt(20) + 1;
                list.add(e);
            }
            return list;
        }
    }

    public static class IndexSheet<T> extends ListSheet<T> {
        // 0: empty 1: List 2： Array
        int type = 0;

        @Override
        public Column[] getHeaderColumns() {
            if (!headerReady) {
                // 将第一行做为头表
                if (data == null || data.isEmpty()) append();
                Object o = getFirst();
                // 无数据输出空
                if (o == null) columns = new Column[0];
                    // List
                else if (List.class.isAssignableFrom(o.getClass())) {
                    type = 1;
                    List row0 = (List) o;
                    columns = new Column[row0.size()];
                    int i = 0;
                    for (Object e : row0) columns[i++] = new Column(e.toString());
                    // 这里取了第一行所以将start向前移动
                    start++;
                }
                // 数组
                else if (o.getClass().isArray()) {
                    type = 2;
                    int len = Array.getLength(o);
                    columns = new Column[len];
                    for (int i = 0; i < len; i++) columns[i] = new Column(Array.get(o, i).toString());
                    // 这里取了第一行所以将start向前移动
                    start++;
                }
            }
            return columns;
        }

        @Override
        protected void resetBlockData() {
            if (!eof && left() < rowBlock.capacity()) append();

            // Find the end index of row-block
            int end = getEndIndex(), len = columns.length;
            for (; start < end; rows++, start++) {
                Row row = rowBlock.next();
                row.index = rows;
                Cell[] cells = row.realloc(len);
                T o = data.get(start);
                boolean isNull = o == null;
                List sub = !isNull && type == 1 ? (List) o : null;
                for (int i = 0; i < len; i++) {
                    // Clear cells
                    Cell cell = cells[i];
                    cell.clear();

                    Object e;
                    Column column = columns[i];
                    if (column.isIgnoreValue() || isNull) e = null;
                        // 根据下标取数
                    else {
                        switch (type) {
                            case 1: e = sub.get(i);      break;
                            case 2: e = Array.get(o, i); break;
                            default: e = null;
                        }
                    }

                    cellValueAndStyle.reset(row, cell, e, column);
                }
                row.height = getRowHeight();
            }
        }
    }

    @Test public void testByIndex() throws IOException {
        List<Object> rows = new ArrayList<>();
        rows.add(Arrays.asList("列1", "列2", "列3"));
        rows.add(Arrays.asList(1, 2, 3));
        rows.add(Arrays.asList(4, 5, 6));
        final String fileName = "by index.xlsx";
        new Workbook().addSheet(new IndexSheet<>().setData(rows)).writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            Iterator<org.ttzero.excel.reader.Row> iter = reader.sheet(0).iterator();
            assertTrue(iter.hasNext());
            org.ttzero.excel.reader.Row row = iter.next();
            assertEquals(row.toString(), "列1 | 列2 | 列3");


            assertTrue(iter.hasNext());
            row = iter.next();
            assertEquals(row.toString(), "1 | 2 | 3");


            assertTrue(iter.hasNext());
            row = iter.next();
            assertEquals(row.toString(), "4 | 5 | 6");
        }
    }

    /**
     * 自定义平铺WorksheetWriter
     */
    public static class TileXMLWorksheetWriter extends XMLWorksheetWriter {
        private int tile; // 平铺的数量，也就是每行重复输出多少条数据
        private String date; // 可忽略，仅仅是表头上的日期

        public TileXMLWorksheetWriter(int tile) {
            this.tile = tile;
        }

        public TileXMLWorksheetWriter(int tile, String date) {
            this.tile = tile;
            this.date = date;
        }

        public int getTile() {
            return tile;
        }

        public void setTile(int tile) {
            this.tile = tile;
        }

        public String getDate() {
            return date;
        }

        public void setDate(String date) {
            this.date = date;
        }

        @Override
        protected void writeBefore() throws IOException {
            // The header columns
            columns = sheet.getAndSortHeaderColumns();
            // Give new columns
            tileColumns();

            boolean nonHeader = sheet.getNonHeader() == 1;

            bw.write(Const.EXCEL_XML_DECLARATION);
            // Declaration
            bw.newLine();
            // Root node
            writeRootNode();

            // Dimension
            writeDimension();

            // SheetViews default value
            writeSheetViews();

            // Default row height and width
            int fillSpace = 6;
            BigDecimal width = BigDecimal.valueOf(!nonHeader ? sheet.getDefaultWidth() : 8.38D);
            String defaultWidth = width.setScale(2, BigDecimal.ROUND_HALF_UP).toString();
            writeSheetFormat();

            // cols
            writeCols(fillSpace, defaultWidth);
        }

        protected void tileColumns() {
            if (tile == 1) return;

            int x = columns.length, y = x * tile, t = columns[columns.length - 1].getRealColIndex();
            // Bound check
            if (y > Const.Limit.MAX_COLUMNS_ON_SHEET)
                throw new TooManyColumnsException(y, Const.Limit.MAX_COLUMNS_ON_SHEET);

            Column[] _columns = new Column[y];
            for (int i = 0; i < y; i++) {
                // 第一个对象的表头不需要复制
                Column col = i < x ? columns[i] : new Column(columns[i % x]).addSubColumn(new Column());
                col.realColIndex = columns[i % x].realColIndex + t * (i / x);
                _columns[i] = col;

                // 替换拣货单上的日期
                Column _col = col;
                do {
                    if (StringUtil.isNotEmpty(_col.getName()) && _col.getName().contains("{date}"))
                        _col.setName(_col.getName().replace("{date}", date));
                }
                while ((_col = _col.next) != null);
            }

            this.columns = _columns;

            // FIXME 这里强行指定合并替换掉原本的头
            List<Dimension> mergeCells = Collections.singletonList(new Dimension(1, (short) 1, 1, (short) y));
            sheet.putExtProp(Const.ExtendPropertyKey.MERGE_CELLS, mergeCells);
        }

        @Override
        protected void writeRow(Row row) throws IOException {
            Cell[] cells = row.getCells();
            int len = cells.length, r = row.getIndex() / tile + startRow, c = columns[columns.length - 1].realColIndex / tile, y = row.getIndex() % tile;
            if (y == 0) startRow(r - startRow, columns[columns.length - 1].realColIndex, -1D);

            // 循环写单元格
            for (int i = row.fc; i < row.lc; i++) writeCell(cells[i], r, i + c * y);

            // 注意这里可能不会关闭row需要在writeAfter进行二次处理
            if (y == tile - 1)
                bw.write("</row>");
        }

        @Override
        protected void writeAfter(int total) throws IOException {
            if (total > 0 && (total - 1) % tile < tile - 1) bw.write("</row>");
            super.writeAfter(total);
        }
    }


    private static org.ttzero.excel.reader.Row createHeaderRow () {
        org.ttzero.excel.reader.Row headerRow = new org.ttzero.excel.reader.Row() {};
        Cell[] cells = new Cell[2];
        cells[0] = new Cell((short) 1).setString("id");
        cells[1] = new Cell((short) 2).setString("name");
        headerRow.setCells(cells);
        return headerRow;
    }


    public static class Template {
        @ExcelColumn(maxWidth = 12.0D, wrapText = true)
        String v1;
        @ExcelColumn(maxWidth = 20.0, wrapText = true)
        String v2;
        @ExcelColumn(maxWidth = 25.0D, wrapText = true)
        String v3;

        static Template of(String v1, String v2, String v3) {
            Template v = new Template();
            v.v1 = v1;
            v.v2 = v2;
            v.v3 = v3;
            return v;
        }

        @Override
        public boolean equals(Object o) {
            if (this == o) return true;
            if (o == null || getClass() != o.getClass()) return false;
            Template template = (Template) o;
            return Objects.equals(v1, template.v1) &&
                Objects.equals(v2, template.v2) &&
                Objects.equals(v3, template.v3);
        }

        @Override
        public int hashCode() {
            return Objects.hash(v1, v2, v3);
        }
    }

    public static class SubModel {
        @ExcelColumn
        private String name;
        @ExcelColumn
        private int status;
    }

    public static class SpecifyConvertModel {
        @ExcelColumn
        private String name;
        @ExcelColumn(converter = StatusConvert.class)
        private int status;

        public static List<SpecifyConvertModel> randomTestData(int n) {
            List<SpecifyConvertModel> list = new ArrayList<>(n);
            for (int i = 0; i < n; i++) {
                SpecifyConvertModel e = new SpecifyConvertModel();
                e.name = getRandomString(10);
                e.status = random.nextInt(4);
                list.add(e);
            }
            return list;
        }

        @Override
        public boolean equals(Object o) {
            if (this == o) return true;
            if (o == null || getClass() != o.getClass()) return false;
            SpecifyConvertModel that = (SpecifyConvertModel) o;
            return status == that.status && Objects.equals(name, that.name);
        }

        @Override
        public int hashCode() {
            return Objects.hash(name, status);
        }
    }

    public static class StatusConvert implements Converter<Integer> {
        final String[] statusDesc = { "未开始", "进行中", "完结", "中止" };

        @Override
        public Integer reversion(String v, Class<?> fieldClazz) {
            for (int i = 0; i < statusDesc.length; i++) {
                if (statusDesc[i].equals(v)) {
                    return i;
                }
            }
            return null;
        }

        @Override
        public Object conversion(Object v) {
            return v != null ? statusDesc[(int) v] : null;
        }
    }

    @Test public void hyperlinkTest() throws IOException {
        List<Item> list = new ArrayList<>();
        list.add(new Item("京东", "https://www.jd.com"));
        list.add(new Item("天猫", "https://www.tmall.com"));
        list.add(new Item("淘宝", "https://www.taobao.com"));

        new Workbook().setAutoSize(true).addSheet(new ListSheet<>(list)).writeTo(defaultTestPath.resolve("超连接测试.xlsx"));
    }

    public static class Item {
        @ExcelColumn
        public String name;
        @Hyperlink
        @ExcelColumn
        public String url;

        public Item(String name, String url) {
            this.name = name;
            this.url = url;
        }
    }

    @Test  public void multipleEnumReversionTest() throws IOException {
        List<MultipleEnumReversionModel> list = MultipleEnumReversionModel.testData();
        new Workbook()
                .addSheet(new ListSheet<>(list))
                .writeTo(defaultTestPath.resolve("multipleEnumReversion.xlsx"));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("multipleEnumReversion.xlsx"))) {
            List<MultipleEnumReversionModel> readList = reader.sheet(0).dataRows().map(row -> row.to(MultipleEnumReversionModel.class)).collect(Collectors.toList());
            assertEquals(list.size(), readList.size());

            MultipleEnumReversionModel multipleEnumReversionModel0 = readList.get(0);
            assertEquals(multipleEnumReversionModel0.getOperator(), Operator.PLUS);
            assertEquals(multipleEnumReversionModel0.getSymbol(), Symbol.PLUS);

            MultipleEnumReversionModel multipleEnumReversionModel1 = readList.get(1);
            assertEquals(multipleEnumReversionModel1.getOperator(), Operator.REDUCE);
            assertEquals(multipleEnumReversionModel1.getSymbol(), Symbol.REDUCE);
        }
    }


    public static class MultipleEnumReversionModel {
        @ExcelColumn(value = "运算符", converter = MultipleEnumConverter.class)
        private Operator operator;
        @ExcelColumn(value = "符号", converter = MultipleEnumConverter.class)
        private Symbol symbol;


        public Operator getOperator() {
            return this.operator;
        }

        public Symbol getSymbol() {
            return this.symbol;
        }

        public void setOperator(Operator operator) {
            this.operator = operator;
        }

        public void setSymbol(Symbol symbol) {
            this.symbol = symbol;
        }

        public static List<MultipleEnumReversionModel> testData() {
            List<MultipleEnumReversionModel> reversionModels = new ArrayList<>();
            MultipleEnumReversionModel model0 = new MultipleEnumReversionModel();
            model0.setOperator(Operator.PLUS);
            model0.setSymbol(Symbol.PLUS);
            reversionModels.add(model0);

            MultipleEnumReversionModel model1 = new MultipleEnumReversionModel();
            model1.setOperator(Operator.REDUCE);
            model1.setSymbol(Symbol.REDUCE);
            reversionModels.add(model1);
            return reversionModels;
        }
    }

    public interface IBaseEnum {
        String caption();
    }

    public enum Operator implements IBaseEnum {

        PLUS("加"),
        REDUCE("减");


        private final String caption;

        Operator(String caption) {
            this.caption = caption;
        }


        @Override
        public String caption() {
            return caption;
        }
    }

    public enum Symbol implements IBaseEnum {

        PLUS("加"),
        REDUCE("减");


        private final String caption;

        Symbol(String caption) {
            this.caption = caption;
        }


        @Override
        public String caption() {
            return caption;
        }
    }

    public static class MultipleEnumConverter implements Converter<IBaseEnum> {
        Map<String, Map<String, IBaseEnum>> map = new HashMap<>(4);

        public MultipleEnumConverter() {
            Map<String, IBaseEnum> operatorMap = new HashMap<>();
            operatorMap.put(Operator.PLUS.caption(), Operator.PLUS);
            operatorMap.put(Operator.REDUCE.caption(), Operator.REDUCE);
            map.put(Operator.class.getSimpleName(), operatorMap);

            Map<String, IBaseEnum> symbolMap = new HashMap<>();
            symbolMap.put(Symbol.PLUS.caption(), Symbol.PLUS);
            symbolMap.put(Symbol.REDUCE.caption(), Symbol.REDUCE);
            map.put(Symbol.class.getSimpleName(), symbolMap);
        }


        @Override
        public IBaseEnum reversion(String v, Class<?> fieldClazz) {

            Map<String, IBaseEnum> clazzIBaseEnumMap = map.get(fieldClazz.getSimpleName());
            if (clazzIBaseEnumMap == null) {
                return null;
            }
            return clazzIBaseEnumMap.get(v);
        }

        @Override
        public Object conversion(Object v) {
            IBaseEnum v1 = (IBaseEnum)  v;
            return v1.caption();
        }
    }

}

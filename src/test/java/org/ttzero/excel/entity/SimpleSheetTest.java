/*
 * Copyright (c) 2017-2024, guanquan.wang@hotmail.com All Rights Reserved.
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
import org.ttzero.excel.entity.e7.XMLCellValueAndStyle;
import org.ttzero.excel.entity.style.Border;
import org.ttzero.excel.entity.style.BorderStyle;
import org.ttzero.excel.entity.style.Fill;
import org.ttzero.excel.entity.style.Font;
import org.ttzero.excel.entity.style.Horizontals;
import org.ttzero.excel.entity.style.PatternType;
import org.ttzero.excel.entity.style.Styles;
import org.ttzero.excel.manager.Const;
import org.ttzero.excel.reader.Cell;
import org.ttzero.excel.reader.Dimension;
import org.ttzero.excel.reader.Drawings;
import org.ttzero.excel.reader.ExcelReader;
import org.ttzero.excel.reader.Row;

import java.awt.Color;
import java.io.IOException;
import java.sql.Timestamp;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.Iterator;
import java.util.List;
import java.util.stream.Collectors;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertTrue;
import static org.ttzero.excel.util.DateUtil.toDateTimeString;

/**
 * @author guanquan.wang
 * @since 2024-09-30
 */
public class SimpleSheetTest extends WorkbookTest {
    @Test public void testConstructor1() throws IOException {
        String fileName = "test simple sheet Constructor1.xlsx";
        new Workbook()
            .setAutoSize(true)
            .addSheet(new SimpleSheet<>(getSimpleDataRows()))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            assertSimpleRows(reader.sheet(0).iterator());
        }
    }

    @Test public void testConstructor2() throws IOException {
        String fileName = "test simple sheet Constructor2.xlsx";
        new Workbook()
            .addSheet(new SimpleSheet<>("Item").setData(getSimpleDataRows()))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            assertEquals("Item", reader.sheet(0).getName());
            assertSimpleRows(reader.sheet(0).iterator());
        }
    }

    @Test public void testConstructor3() throws IOException {
        String fileName = "test simple sheet Constructor3.xlsx";
        new Workbook()
            .addSheet(new SimpleSheet<>("Item"
                , new Column("ID", "id")
                , new Column("NAME", "name"))
                .setData(getSimpleDataRows()))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            assertEquals("Item", reader.sheet(0).getName());
            assertSimpleRows(reader.sheet(0).iterator(), "ID | NAME");
        }
    }

    @Test public void testConstructor4() throws IOException {
        String fileName = "test simple sheet Constructor4.xlsx";
        new Workbook()
            .setAutoSize(true)
            .addSheet(new SimpleSheet<>("Item", WaterMark.of(author)
                , new Column("ID", "id").setFont(new Font("华文行楷", 24, Color.BLUE))
                , new Column("NAME", "name").setFont(new Font("微软雅黑", 18)))
                .setData(getSimpleDataRows()))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            org.ttzero.excel.reader.Sheet sheet = reader.sheet(0);
            assertEquals("Item", sheet.getName());
            assertSimpleRows(reader.sheet(0).iterator(), "ID | NAME");

            List<Drawings.Picture> pictures = sheet.listPictures();
            assertEquals(pictures.size(), 1);
            Drawings.Picture pic = pictures.get(0);
            assertTrue(pic.isBackground());
        }
    }

    @Test public void testConstructor5() throws IOException {
        String fileName = "test simple sheet Constructor5.xlsx";
        new Workbook()
            .setAutoSize(true)
            .addSheet(new SimpleSheet<>(getSimpleDataRows()))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            assertSimpleRows(reader.sheet(0).iterator());
        }
    }


    @Test public void testConstructor9() throws IOException {
        String fileName = "test simple sheet Constructor9.xlsx";
        new Workbook()
            .setAutoSize(true)
            .addSheet(new SimpleSheet<>(getSimpleDataRows()
                , WaterMark.of(author)
                , new Column("ID", "id")
                , new Column("NAME", "name")))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            assertSimpleRows(reader.sheet(0).iterator(), "ID | NAME");
        }
    }

    @Test public void testConstructor10() throws IOException {
        String fileName = "test simple sheet Constructor10.xlsx";
        new Workbook()
            .setAutoSize(true)
            .addSheet(new SimpleSheet<>("ITEM"
                , getSimpleDataRows()
                , WaterMark.of(author)
                , new Column("ID", "id")
                , new Column("NAME", "name")))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            org.ttzero.excel.reader.Sheet sheet = reader.sheet(0);
            assertEquals("ITEM", sheet.getName());
            assertSimpleRows(reader.sheet(0).iterator(), "ID | NAME");
            List<Drawings.Picture> pictures = sheet.listPictures();
            assertEquals(pictures.size(), 1);
            Drawings.Picture pic = pictures.get(0);
            assertTrue(pic.isBackground());
        }
    }

    @Test public void testConstructor11() throws IOException {
        String fileName = "test simple sheet Constructor11.xlsx";
        new Workbook()
            .setAutoSize(true)
            .addSheet(new SimpleSheet<>(new Column("ID", "id")
                , new Column("NAME", "name")))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            org.ttzero.excel.reader.Sheet sheet = reader.sheet(0);
            assertEquals("Sheet1", sheet.getName());
            assertEquals(sheet.rows().count(), 1L);
        }
    }

    @Test public void testSimpleSheet() throws IOException {
        final String fileName = "list simple sheet.xlsx";
        new Workbook().addSheet(new SimpleSheet<>(getSimpleDataRows())).writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            assertSimpleRows(reader.sheet(0).iterator());
        }
    }

    @Test public void testSimpleSheetFirstRowAsHeader() throws IOException {
        final String fileName2 = "list simple sheet - first-row-as-header.xlsx";
        new Workbook().addSheet(new SimpleSheet<>(getSimpleDataRows()).firstRowAsHeader()).writeTo(defaultTestPath.resolve(fileName2));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName2))) {
            assertSimpleRows(reader.sheet(0).iterator());
        }
    }

    @Test public void testSimpleSheetSpecifyHeader() throws IOException {
        Date now = new Date();
        List<Object> rows = new ArrayList<>();
        rows.add(new String[]{"列1", "列2", "列3"});
        rows.add(new int[]{1, 2, 3, 4});
        rows.add(new Object[]{5, now, 7, null, "字母", 9, 10.1243});
        final String fileName3 = "list simple sheet - specify header.xlsx";

        new Workbook()
            .addSheet(new SimpleSheet<>(rows).setHeader(Arrays.asList("表头1", "表头2")))
            .writeTo(defaultTestPath.resolve(fileName3));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName3))) {
            Iterator<org.ttzero.excel.reader.Row> iter = reader.sheet(0).iterator();
            assertTrue(iter.hasNext());
            org.ttzero.excel.reader.Row row = iter.next();
            assertEquals("表头1 | 表头2", row.toString());
            Styles styles = reader.getStyles();
            int fx = row.getCellStyle(0);
            Font font = styles.getFont(fx);
            assertEquals(new Font("宋体", 12, Font.Style.BOLD, Color.BLACK), font);
            Fill fill = styles.getFill(fx);
            assertEquals(new Fill(Styles.toColor("#E9EAEC")), fill);

            assertTrue(iter.hasNext());
            row = iter.next();
            assertEquals("列1 | 列2 | 列3", row.toString());

            assertTrue(iter.hasNext());
            row = iter.next();
            assertEquals("1 | 2 | 3 | 4", row.toString());

            assertTrue(iter.hasNext());
            row = iter.next();
            // 时间忽略毫秒值，所以这里特殊处理
            assertEquals("5 | " + toDateTimeString(now) + " | 7 | null | 字母 | 9 | 10.1243", row.toString());
        }
    }

    @Test public void testSimpleSheetPutObject() throws IOException {
        List<ListObjectSheetTest.Student> expectList = ListObjectSheetTest.Student.randomTestData();
        final String fileName = "list simple sheet put objects.xlsx";
        new Workbook().addSheet(new SimpleSheet<>(expectList)).writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            List<ListObjectSheetTest.Student> list = reader.sheet(0).dataRows().map(row -> row.to(ListObjectSheetTest.Student.class)).collect(Collectors.toList());
            assertEquals(expectList.size(), list.size());
            for (int i = 0, len = expectList.size(); i < len; i++) {
                ListObjectSheetTest.Student expect = expectList.get(i), e = list.get(i);
                assertEquals(expect.getName(), e.getName());
                assertEquals(expect.getScore(), e.getScore());
            }
        }
    }

    @Test public void testSimpleSheetPaging() throws IOException {
        final int max = 100;
        List<Object> expectList = new ArrayList<>(max);
        final String fileName = "list simple sheet paging.xlsx";
        new Workbook().setAutoSize(true)
            .addSheet(new SimpleSheet<>().setHeader(Arrays.asList("HEAD1", "HEAD2", "HEAD3")).setData((i, a) -> {
            if (i < max) {
                List<Object> list = new ArrayList<>(10);
                for (int p = 1; p <= 10; p++) {
                    list.add(Arrays.asList(i + p, getRandomAssicString(15), new Timestamp(System.currentTimeMillis() - random.nextInt(5000 + 100))));
                }
                expectList.addAll(list);
                return list;
            }
            return null;
        }).setStartCoordinate(5)).writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            Iterator<Row> iter = reader.sheet(0).iterator();
            // 表头行
            assertTrue(iter.hasNext());
            Row header = iter.next();
            assertEquals(header.getRowNum(), 5);
            assertEquals("HEAD1 | HEAD2 | HEAD3", header.toString());

            // 数据行
            for (Object o : expectList) {
                assertTrue(iter.hasNext());
                @SuppressWarnings("unchecked")
                List<Object> sub = (List<Object>) o;
                Row row = iter.next();
                assertEquals(row.getInt(0), sub.get(0));
                assertEquals(row.getString(1), sub.get(1));
                assertEquals(row.getTimestamp(2).getTime() / 1000, ((Timestamp) sub.get(2)).getTime() / 1000);
            }
        }
    }

    @Test public void t() throws IOException {
        List<Object> rows = new ArrayList<>();
        List<ListObjectSheetTest.Student> students1, students2;
        students1 = new ArrayList<>(ListObjectSheetTest.Student.randomTestData(3));
        students2 = new ArrayList<>(ListObjectSheetTest.Student.randomTestData(3));
        List<Dimension> mergeCells = new ArrayList<>();
        // 循环添加多个对象
        rows.add(new Object[]{"班级1"});
        mergeCells.add(new Dimension(1, (short) 1, 1, (short) 2));
        rows.add(new Object[]{"学号", "姓名"});
        for (ListObjectSheetTest.Student e : students1) {
            rows.add(new Object[]{e.getId(), e.getName()});
        }
        rows.add(new Object[]{});
        rows.add(new Object[]{"班级2"});
        int row = students1.size() + 4;
        mergeCells.add(new Dimension(row, (short) 1, row, (short) 2));
        rows.add(new Object[]{"学号", "姓名"});
        for (ListObjectSheetTest.Student e : students2) {
            rows.add(new Object[]{e.getId(), e.getName()});
        }

        new Workbook().addSheet(new SimpleSheet<>(rows, new Column().autoSize(), new Column().setWidth(20)).setCellValueAndStyle(new XMLCellValueAndStyle() {
            @Override
            public void reset(org.ttzero.excel.entity.Row row, Cell cell, Object e, Column hc) {
                // 将值转输出需要的统一格式
                setCellValue(row, cell, e, hc, hc.getClazz(), hc.getConversion() != null);
                // 可以根据行号和列号
                int xf;
                if (row.index == 0) {
                    int style = 0;
                    Styles styles = hc.styles;
                    style = styles.modifyFont(style, new Font("宋体", 13));
                    style = styles.modifyFill(style, new Fill(PatternType.solid, Color.pink));
                    style = styles.modifyBorder(style, new Border(BorderStyle.THIN, Color.orange));
                    style = styles.modifyHorizontal(style, Horizontals.CENTER);
                    xf = styles.of(style);
                } else {
                    xf = getStyleIndex(row, hc, e);
                }
                cell.xf = xf;
            }
        }).putExtProp(Const.ExtendPropertyKey.MERGE_CELLS, mergeCells)).writeTo(defaultTestPath.resolve("1.xlsx"));
    }

    @Test public void testSpecifyCoordinateWrite() throws IOException {
        final String fileName = "test specify coordinate D4 SimpleSheet.xlsx";
        List<Object> expectList = getSimpleDataRows();
        new Workbook().setAutoSize(true)
            .addSheet(new SimpleSheet<>(expectList).setStartCoordinate("D4"))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            Iterator<org.ttzero.excel.reader.Row> iter = reader.sheet(0).iterator();
            org.ttzero.excel.reader.Row firstRow = iter.next();
            assertNotNull(firstRow);
            assertEquals(firstRow.getRowNum(), 4);
//            assertEquals(firstRow.getFirstColumnIndex(), 3);
            System.out.println(firstRow);
            for (; iter.hasNext();) {
                System.out.println(iter.next());
            }
        }
    }

    private static List<Object> getSimpleDataRows() {
        List<Object> rows = new ArrayList<>();
        rows.add(Arrays.asList("列1", "列2", "列3"));
        rows.add(Arrays.asList(1, 2, 3, 4));
        rows.add(Arrays.asList(5, 6, 7, null, "字母", 9, 10));
        return rows;
    }

    private static void assertSimpleRows(Iterator<Row> iter) {
        assertSimpleRows(iter, null);
    }

    private static void assertSimpleRows(Iterator<Row> iter, String header) {
        assertTrue(iter.hasNext());
        org.ttzero.excel.reader.Row row = iter.next();
        if (header != null) {
            assertEquals(row.toString(), header);
            assertTrue(iter.hasNext());
            row = iter.next();
        }
        assertEquals(row.toString(), "列1 | 列2 | 列3");
//        Styles styles = row.getStyles();
//        int fx = row.getCellStyle(0);
//        Font font = styles.getFont(fx);
//        assertEquals(font, new Font("宋体", 11, Color.BLACK));
//        Fill fill = styles.getFill(fx);
//        assertEquals(fill, new Fill(Styles.toColor("#E9EAEC")));

        assertTrue(iter.hasNext());
        row = iter.next();
        assertEquals(row.toString(), "1 | 2 | 3 | 4");

        assertTrue(iter.hasNext());
        row = iter.next();
        assertEquals(row.toString(), "5 | 6 | 7 | null | 字母 | 9 | 10");
    }
}

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
import org.ttzero.excel.entity.style.Fill;
import org.ttzero.excel.entity.style.Font;
import org.ttzero.excel.entity.style.Styles;
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
import static org.junit.Assert.assertTrue;

/**
 * @author guanquan.wang
 * @since 2024-09-30
 */
public class SimpleSheetTest extends WorkbookTest {
    @Test public void testConstructor1() throws IOException {
        String fileName = "test simple sheet Constructor1.xlsx";
        new Workbook()
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
                , new Column("ID", "id")
                , new Column("NAME", "name"))
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
            assertEquals(row.toString(), "表头1 | 表头2");
            Styles styles = reader.getStyles();
            int fx = row.getCellStyle(0);
            Font font = styles.getFont(fx);
            assertEquals(font, new Font("宋体", 12, Font.Style.BOLD, Color.BLACK));
            Fill fill = styles.getFill(fx);
            assertEquals(fill, new Fill(Styles.toColor("#E9EAEC")));

            assertTrue(iter.hasNext());
            row = iter.next();
            assertEquals(row.toString(), "列1 | 列2 | 列3");

            assertTrue(iter.hasNext());
            row = iter.next();
            assertEquals(row.toString(), "1 | 2 | 3 | 4");

            assertTrue(iter.hasNext());
            row = iter.next();
            // 时间忽略毫秒值，所以这里特殊处理
            assertEquals(row.toString(), "5 | " + new Timestamp(now.getTime() / 1000 * 1000) + " | 7 | null | 字母 | 9 | 10.1243");
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
        }).setStartRowIndex(5, false)).writeTo(defaultTestPath.resolve(fileName));

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

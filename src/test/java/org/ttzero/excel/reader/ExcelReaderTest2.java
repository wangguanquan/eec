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


package org.ttzero.excel.reader;

import org.junit.Ignore;
import org.junit.Test;
import org.ttzero.excel.annotation.ExcelColumn;
import org.ttzero.excel.entity.Column;
import org.ttzero.excel.entity.EmptySheet;
import org.ttzero.excel.entity.ListMapSheet;
import org.ttzero.excel.entity.ListObjectSheetTest;
import org.ttzero.excel.entity.ListSheet;
import org.ttzero.excel.entity.Panes;
import org.ttzero.excel.entity.Workbook;
import org.ttzero.excel.util.CSVUtil;
import org.ttzero.excel.util.DateUtil;
import org.ttzero.excel.util.StringUtil;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.stream.Collectors;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertTrue;
import static org.ttzero.excel.entity.WorkbookTest.defaultTestPath;
import static org.ttzero.excel.entity.WorkbookTest.getRandomString;
import static org.ttzero.excel.entity.WorkbookTest.random;
import static org.ttzero.excel.reader.ExcelReaderTest.testResourceRoot;

/**
 * @author guanquan.wang at 2023-01-06 09:32
 */
public class ExcelReaderTest2 {
    @Test public void testIsBlank() throws IOException {
        try (ExcelReader reader = ExcelReader.read(testResourceRoot().resolve("#150.xlsx"))) {
            reader.sheet(0).rows().forEach(row -> {
                switch (row.getRowNum()) {
                    case 1:
                        assertFalse(row.isEmpty());
                        assertFalse(row.isBlank());
                        break;
                    case 2:
                        assertFalse(row.isEmpty());
                        assertTrue(row.isBlank());
                        break;
                }
            });
        }
    }

    @Test public void test354() throws IOException {
        try (ExcelReader reader = ExcelReader.read(testResourceRoot().resolve("#354.xlsx"))) {
            for (int i = 0, len = reader.getSheetCount(); i < len; i++) {
                Sheet sheet = reader.sheet(i);
                Path expectPath = testResourceRoot().resolve("expect/#354$" + sheet.getName() + ".txt");
                if (Files.exists(expectPath)) {
                    List<String[]> expectList = CSVUtil.read(expectPath);
                    Iterator<Row> it = sheet.iterator();
                    for (String[] expect : expectList) {
                        assertTrue(it.hasNext());
                        Row row = it.next();

                        for (int start = row.getFirstColumnIndex(), end = row.getLastColumnIndex(); start < end; start++) {
                            Cell cell = row.getCell(start);
                            CellType type = row.getCellType(cell);
                            String e = expect[start], o;
                            if (type == CellType.INTEGER) o = row.getInt(cell).toString();
                            else o = row.getString(start);
                            if (StringUtil.isEmpty(e)) assertTrue(StringUtil.isEmpty(o));
                            else assertEquals(o, e);
                        }
                    }
                } else {
                    for (Iterator<Row> iter = sheet.iterator(); iter.hasNext(); ) {
                        Row row = iter.next();
                        assertTrue(StringUtil.isNotEmpty(row.toString()));
                    }
                }
            }
        }
    }

    @Test public void testMerge() throws IOException {
        try (ExcelReader reader = ExcelReader.read(testResourceRoot().resolve("merge.xlsx"))) {
            MergeSheet sheet = reader.sheet(0).asMergeSheet();
            List<Dimension> list = sheet.getMergeCells();
            assertEquals(list.size(), 4);
            assertEquals(list.get(0), Dimension.of("B2:C2"));
            assertEquals(list.get(1), Dimension.of("E5:F8"));
            assertEquals(list.get(2), Dimension.of("A13:A20"));
            assertEquals(list.get(3), Dimension.of("B16:E17"));

            sheet = reader.sheet(1).asMergeSheet();
            list = sheet.getMergeCells();
            assertEquals(list.size(), 2);
            assertEquals(list.get(0), Dimension.of("BM2:BQ11"));
            assertEquals(list.get(1), Dimension.of("A1:B26"));

            sheet = reader.sheet(2).asMergeSheet();
            list = sheet.getMergeCells();
            assertEquals(list.size(), 2);
            assertEquals(list.get(0), Dimension.of("A16428:D16437"));
            assertEquals(list.get(1), Dimension.of("A1:K3"));

            sheet = reader.sheet(3).asMergeSheet();
            list = sheet.getMergeCells();
            assertEquals(list.size(), 1);
            assertEquals(list.get(0), Dimension.of("A1:CF1434"));
        }
    }

    @Test public void testLargeMerge() throws IOException {
        try (ExcelReader reader = ExcelReader.read(testResourceRoot().resolve("largeMerged.xlsx"))) {
            MergeSheet sheet = reader.sheet(0).asMergeSheet();
            List<Dimension> list = sheet.getMergeCells();
            assertEquals(list.size(), 2608);
            assertEquals(list.get(0), Dimension.of("C3:F3"));
            assertEquals(list.get(1), Dimension.of("J2:J3"));
            assertEquals(list.get(2), Dimension.of("B2:B3"));
            assertEquals(list.get(3), Dimension.of("C5:F5"));

            assertEquals(list.get(98), Dimension.of("C82:F82"));
            assertEquals(list.get(120), Dimension.of("A104:A106"));
            assertEquals(list.get(210), Dimension.of("C176:F176"));
            assertEquals(list.get(984), Dimension.of("C821:F821"));

            assertEquals(list.get(1626), Dimension.of("B1362:B1371"));
            assertEquals(list.get(1627), Dimension.of("J1362:J1363"));
            assertEquals(list.get(2381), Dimension.of("B2006:B2007"));
            assertEquals(list.get(2396), Dimension.of("J2019:J2020"));

            assertEquals(list.get(2596), Dimension.of("C2190:F2190"));
            assertEquals(list.get(2601), Dimension.of("J2195:J2196"));
            assertEquals(list.get(2605), Dimension.of("C2198:F2198"));
            assertEquals(list.get(2607), Dimension.of("C2200:F2200"));
        }
    }

    @Test public void testForceImport() throws IOException {
        Map<String, Object> data1 = new HashMap<>();
        data1.put("id", 1);
        data1.put("name", "abc");

        Map<String, Object> data2 = new HashMap<>();
        data2.put("id", 2);
        data2.put("name", "xyz");
        new Workbook()
            .addSheet(new ListMapSheet().setData(Arrays.asList(data1, data2)))
            .writeTo(defaultTestPath.resolve("Force Import.xlsx"));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("Force Import.xlsx"))) {
            List<U> list = reader.sheet(0).forceImport().dataRows().map(row -> row.to(U.class)).collect(Collectors.toList());
            assertEquals(list.size(), 2);
            assertEquals("1: abc", list.get(0).toString());
            assertEquals("2: xyz", list.get(1).toString());
        }
    }

    @Test public void testUpperCaseRead() throws IOException {
        Map<String, Object> data1 = new HashMap<>();
        data1.put("ID", 1);
        data1.put("NAME", "abc");

        Map<String, Object> data2 = new HashMap<>();
        data2.put("ID", 2);
        data2.put("NAME", "xyz");

        new Workbook()
            .addSheet(new ListMapSheet(Arrays.asList(data1, data2)))
            .writeTo(defaultTestPath.resolve("Upper case Reader test.xlsx"));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("Upper case Reader test.xlsx"))) {
            List<U> list = reader.sheet(0).forceImport().dataRows().map(row -> row.to(U.class)).collect(Collectors.toList());
            assertEquals(list.size(), 2);
            assertEquals("0: null", list.get(0).toString());
            assertEquals("0: null", list.get(1).toString());

            list = reader.sheet(0).reset().addHeaderColumnReadOption(HeaderRow.FORCE_IMPORT | HeaderRow.IGNORE_CASE)
                .dataRows().map(row -> row.to(U.class)).collect(Collectors.toList());
            assertEquals(list.size(), 2);
            assertEquals("1: abc", list.get(0).toString());
            assertEquals("2: xyz", list.get(1).toString());
        }
    }

    @Test public void testCamelCaseRead() throws IOException {
        Map<String, Object> data1 = new HashMap<>();
        data1.put("USER_ID", 1);
        data1.put("USER_NAME", "abc");

        Map<String, Object> data2 = new HashMap<>();
        data2.put("USER_ID", 2);
        data2.put("USER_NAME", "xyz");

        new Workbook()
            .addSheet(new ListMapSheet(Arrays.asList(data1, data2)))
            .writeTo(defaultTestPath.resolve("Underline case Reader test.xlsx"));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("Underline case Reader test.xlsx"))) {
            List<User> list = reader.sheet(0).forceImport().dataRows().map(row -> row.to(User.class)).collect(Collectors.toList());
            assertEquals(list.size(), 2);
            assertEquals("0: null", list.get(0).toString());
            assertEquals("0: null", list.get(1).toString());

            list = reader.sheet(0).reset().addHeaderColumnReadOption(HeaderRow.FORCE_IMPORT | HeaderRow.CAMEL_CASE)
                .dataRows().map(row -> row.to(User.class)).collect(Collectors.toList());
            assertEquals(list.size(), 2);
            assertEquals("1: abc", list.get(0).toString());
            assertEquals("2: xyz", list.get(1).toString());
        }
    }

    @Test public void testEmptyBindObj() throws IOException {
        new Workbook().addSheet(new EmptySheet()).writeTo(defaultTestPath.resolve("empty.xlsx"));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("empty.xlsx"))) {
            Sheet sheet = reader.sheet(0);
            List<U> list = sheet.header(1, 2).bind(U.class).rows().map(row -> (U) row.get()).collect(Collectors.toList());
            assertTrue(list.isEmpty());

            list = sheet.reset().header(1).rows().map(row -> row.to(U.class)).collect(Collectors.toList());
            assertTrue(list.isEmpty());

            list = sheet.reset().dataRows().map(row -> row.to(U.class)).collect(Collectors.toList());
            assertTrue(list.isEmpty());
        }
    }

    @Ignore
    @Test public void test200w() throws IOException {
        final int row_len = 2_000_000;
        // 0: 写入数据总行数
        // 1: nv大于1w的行数
        final int[] expect = { 0, 0 };
        new Workbook()
            .onProgress((sheet, rows) -> System.out.println(sheet.getName() + " 已写入: " + rows))
            .addSheet(new ListSheet<E>().setData((i, lastOne) -> {
                List<E> list = null;
                if (i < row_len) {
                    list = E.data();
                    expect[0] += list.size();
                    expect[1] += list.stream().filter(e -> e.nv > 10000).count();
                }
                return list;
            }))
            .writeTo(defaultTestPath.resolve("200w.xlsx"));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("200w.xlsx"))) {
            long count = reader.sheets().flatMap(Sheet::dataRows).count();
            assertEquals(count, expect[0]);
            long count1w = reader.sheets().flatMap(Sheet::dataRows).map(row -> row.getInt(0)).filter(i -> i > 10000).count();
            assertEquals(count1w, expect[1]);
        }
    }

    @Test public void testEntryMissKey() throws IOException {
        List<ListObjectSheetTest.Item> expectList = ListObjectSheetTest.Item.randomTestData(10);
        new Workbook().addSheet(new ListSheet<ListObjectSheetTest.Item>(
                new Column("id"), new Column("name"))
                .setData(expectList))
            .writeTo(defaultTestPath.resolve("test entry miss key.xlsx"));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("test entry miss key.xlsx"))) {
            List<ListObjectSheetTest.Item> list = reader.sheet(0).dataRows().map(row -> row.to(ListObjectSheetTest.Item.class)).collect(Collectors.toList());
            assertTrue(listEquals(list, expectList));
        }
    }

    @Test public void testFullReader() throws IOException {
        try (ExcelReader reader = ExcelReader.read(testResourceRoot().resolve("1.xlsx"))) {
            FullSheet sheet = reader.sheet(0).asFullSheet();
            Panes panes = sheet.getFreezePanes();
            assertEquals(panes.row, 1);
            assertEquals(panes.col, 0);
            assertTrue(sheet.isShowGridLines());
            assertEquals(Dimension.of("A1:F1"), sheet.getFilter());
            List<Col> list = sheet.getCols();
            assertEquals(list.size(), 6);
            assertTrue(list.get(2).hidden);
            assertEquals((int) sheet.getDefaultColWidth(), 21);
            assertEquals((int) sheet.getDefaultRowHeight(), 15);

//            List<Dimension> mergeCells = sheet.getMergeCells();
//            assertEquals(mergeCells.size(), 1);
//            assertEquals(Dimension.of("B98:E100"), mergeCells.get(0));

            Path expectPath = testResourceRoot().resolve("expect/1$" + sheet.getName() + ".txt");
            if (Files.exists(expectPath)) {
                List<String[]> expectList = CSVUtil.read(expectPath);

                Iterator<Row> it = sheet.iterator();
                for (String[] expect : expectList) {
                    assertTrue(it.hasNext());
                    Row row = it.next();

                    // 第20行隐藏
                    if (row.getRowNum() == 20) assertTrue(row.isHidden());

                    for (int start = row.getFirstColumnIndex(), end = row.getLastColumnIndex(); start < end; start++) {
                        Cell cell = row.getCell(start);
                        CellType type = row.getCellType(cell);
                        String e = expect[start], o;
                        switch (type) {
                            case INTEGER : o = row.getInt(cell).toString();                   break;
                            case BOOLEAN : o = row.getBoolean(cell).toString().toUpperCase(); break;
                            case DATE    : o = DateUtil.toString(row.getDate(cell));          break;
                            default: o = row.getString(start);
                        }
                        if (StringUtil.isEmpty(e)) assertTrue(StringUtil.isEmpty(o));
                        else assertEquals(o, e);
                    }
                }
            }
        }
    }

    public static <T> boolean listEquals(List<T> list, List<T> expectList) {
        if (list == expectList) return true;
        if (list == null || expectList == null) return false;

        int length = list.size(), i = 0;
        if (expectList.size() != length)
            return false;

        for (; i < length && Objects.equals(list.get(i), expectList.get(i)); i++);

        return i == length;
    }

    public static class U {
        int id;
        String name;

        public void setId(int id) {
            this.id = id;
        }

        @Override
        public String toString() {
            return id + ": " + name;
        }
    }

    public static class User {
        int userId;
        String userName;

        @Override
        public String toString() {
            return userId + ": " + userName;
        }
    }

    public static class E {
        @ExcelColumn
        private int nv;
        @ExcelColumn
        private String str;

        public static List<E> data() {
            List<E> list = new ArrayList<>(1000);
            for (int i = 0; i < 1000; i++) {
                E e = new E();
                list.add(e);
                e.nv = random.nextInt();
                e.str = getRandomString();
            }
            return list;
        }

    }
}

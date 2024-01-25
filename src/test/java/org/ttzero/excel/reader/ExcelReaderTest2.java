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
                        assert !row.isEmpty();
                        assert !row.isBlank();
                        break;
                    case 2:
                        assert !row.isEmpty();
                        assert row.isBlank();
                        break;
                }
            });
        }
    }

    @Test public void test354() throws IOException {
        try (ExcelReader reader = ExcelReader.read(testResourceRoot().resolve("#354.xlsx"))) {
            for (int i = 0, len = reader.getSize(); i < len; i++) {
                Sheet sheet = reader.sheet(i);
                Path expectPath = testResourceRoot().resolve("expect/#354$" + sheet.getName() + ".txt");
                if (Files.exists(expectPath)) {
                    List<String[]> expectList = CSVUtil.read(expectPath);
                    Iterator<Row> it = sheet.iterator();
                    for (String[] expect : expectList) {
                        assert it.hasNext();
                        Row row = it.next();

                        for (int start = row.getFirstColumnIndex(), end = row.getLastColumnIndex(); start < end; start++) {
                            Cell cell = row.getCell(start);
                            CellType type = row.getCellType(cell);
                            String e = expect[start], o;
                            if (type == CellType.INTEGER) o = row.getInt(cell).toString();
                            else o = row.getString(start);
                            assert StringUtil.isEmpty(e) && StringUtil.isEmpty(o) || e.equals(o);
                        }
                    }
                } else {
                    for (Iterator<Row> iter = sheet.iterator(); iter.hasNext(); ) {
                        Row row = iter.next();
                        assert  StringUtil.isNotEmpty(row.toString());
                    }
                }
            }
        }
    }

    @Test public void testMerge() throws IOException {
        try (ExcelReader reader = ExcelReader.read(testResourceRoot().resolve("merge.xlsx"))) {
            MergeSheet sheet = reader.sheet(0).asMergeSheet();
            List<Dimension> list = sheet.getMergeCells();
            assert list.size() == 4;
            assert list.get(0).equals(Dimension.of("B2:C2"));
            assert list.get(1).equals(Dimension.of("E5:F8"));
            assert list.get(2).equals(Dimension.of("A13:A20"));
            assert list.get(3).equals(Dimension.of("B16:E17"));

            sheet = reader.sheet(1).asMergeSheet();
            list = sheet.getMergeCells();
            assert list.size() == 2;
            assert list.get(0).equals(Dimension.of("BM2:BQ11"));
            assert list.get(1).equals(Dimension.of("A1:B26"));

            sheet = reader.sheet(2).asMergeSheet();
            list = sheet.getMergeCells();
            assert list.size() == 2;
            assert list.get(0).equals(Dimension.of("A16428:D16437"));
            assert list.get(1).equals(Dimension.of("A1:K3"));

            sheet = reader.sheet(3).asMergeSheet();
            list = sheet.getMergeCells();
            assert list.size() == 1;
            assert list.get(0).equals(Dimension.of("A1:CF1434"));
        }
    }

    @Test public void testLargeMerge() throws IOException {
        try (ExcelReader reader = ExcelReader.read(testResourceRoot().resolve("largeMerged.xlsx"))) {
            MergeSheet sheet = reader.sheet(0).asMergeSheet();
            List<Dimension> list = sheet.getMergeCells();
            assert list.size() == 2608;
            assert list.get(0).equals(Dimension.of("C3:F3"));
            assert list.get(1).equals(Dimension.of("J2:J3"));
            assert list.get(2).equals(Dimension.of("B2:B3"));
            assert list.get(3).equals(Dimension.of("C5:F5"));

            assert list.get(98).equals(Dimension.of("C82:F82"));
            assert list.get(120).equals(Dimension.of("A104:A106"));
            assert list.get(210).equals(Dimension.of("C176:F176"));
            assert list.get(984).equals(Dimension.of("C821:F821"));

            assert list.get(1626).equals(Dimension.of("B1362:B1371"));
            assert list.get(1627).equals(Dimension.of("J1362:J1363"));
            assert list.get(2381).equals(Dimension.of("B2006:B2007"));
            assert list.get(2396).equals(Dimension.of("J2019:J2020"));

            assert list.get(2596).equals(Dimension.of("C2190:F2190"));
            assert list.get(2601).equals(Dimension.of("J2195:J2196"));
            assert list.get(2605).equals(Dimension.of("C2198:F2198"));
            assert list.get(2607).equals(Dimension.of("C2200:F2200"));
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
            assert list.size() == 2;
            assert "1: abc".equals(list.get(0).toString());
            assert "2: xyz".equals(list.get(1).toString());
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
            assert list.size() == 2;
            assert "0: null".equals(list.get(0).toString());
            assert "0: null".equals(list.get(1).toString());

            list = reader.sheet(0).reset().addHeaderColumnReadOption(HeaderRow.FORCE_IMPORT | HeaderRow.IGNORE_CASE)
                .dataRows().map(row -> row.to(U.class)).collect(Collectors.toList());
            assert list.size() == 2;
            assert "1: abc".equals(list.get(0).toString());
            assert "2: xyz".equals(list.get(1).toString());
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
            assert list.size() == 2;
            assert "0: null".equals(list.get(0).toString());
            assert "0: null".equals(list.get(1).toString());

            list = reader.sheet(0).reset().addHeaderColumnReadOption(HeaderRow.FORCE_IMPORT | HeaderRow.CAMEL_CASE)
                .dataRows().map(row -> row.to(User.class)).collect(Collectors.toList());
            assert list.size() == 2;
            assert "1: abc".equals(list.get(0).toString());
            assert "2: xyz".equals(list.get(1).toString());
        }
    }

    @Test public void testEmptyBindObj() throws IOException {
        new Workbook().addSheet(new EmptySheet()).writeTo(defaultTestPath.resolve("empty.xlsx"));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("empty.xlsx"))) {
            Sheet sheet = reader.sheet(0);
            List<U> list = sheet.header(1, 2).bind(U.class).rows().map(row -> (U) row.get()).collect(Collectors.toList());
            assert list.isEmpty();

            list = sheet.reset().header(1).rows().map(row -> row.to(U.class)).collect(Collectors.toList());
            assert list.isEmpty();

            list = sheet.reset().dataRows().map(row -> row.to(U.class)).collect(Collectors.toList());
            assert list.isEmpty();
        }
    }

    @Ignore
    @Test public void testFullSheet() throws IOException {
        final int loop = 2000;
        new Workbook("200w").addSheet(new ListSheet<E>() {
            int n = 0; // 页码
            @Override
            public List<E> more() {
                return n++ < loop ? E.data() : null;
            }
        }).writeTo(defaultTestPath);
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
            assert panes.row == 1;
            assert panes.col == 0;
            assert sheet.showGridLines();
            assert sheet.getFilter().equals(Dimension.of("A1:F1"));
            List<Col> list = sheet.getCols();
            assert list.size() == 6;
            assert list.get(2).hidden;
            assert (int) sheet.getDefaultColWidth() == 30;
            assert (int) sheet.getDefaultRowHeight() == 15;

            List<Dimension> mergeCells = sheet.getMergeCells();
            assert mergeCells.size() == 1;
            assert mergeCells.get(0).equals(Dimension.of("G3:H6"));
            sheet.forceImport().dataRows().map(row -> row.to(ExcelReaderTest.AnnotationEntry.class)).forEach(System.out::println);
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

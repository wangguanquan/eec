/*
 * Copyright (c) 2017-2023, guanquan.wang@hotmail.com All Rights Reserved.
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

import org.junit.Test;
import org.ttzero.excel.entity.Column;
import org.ttzero.excel.entity.Comment;
import org.ttzero.excel.entity.ListMapSheet;
import org.ttzero.excel.entity.ListObjectSheetTest;
import org.ttzero.excel.entity.ListSheet;
import org.ttzero.excel.entity.Panes;
import org.ttzero.excel.entity.Relationship;
import org.ttzero.excel.entity.Workbook;
import org.ttzero.excel.manager.Const;
import org.ttzero.excel.manager.RelManager;
import org.ttzero.excel.util.CSVUtil;
import org.ttzero.excel.util.StringUtil;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.sql.Timestamp;
import java.time.LocalDate;
import java.util.Arrays;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.stream.Collectors;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertTrue;
import static org.ttzero.excel.entity.WorkbookTest.defaultTestPath;
import static org.ttzero.excel.reader.ExcelReaderTest.testResourceRoot;
import static org.ttzero.excel.reader.ExcelReader.coordinateToLong;
import static org.ttzero.excel.util.DateUtil.toDateTimeString;

/**
 * @author guanquan.wang at 2023-01-06 09:32
 */
public class ExcelReader2Test {
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
            FullSheet sheet = reader.sheet(0).asFullSheet();
            List<Dimension> list = sheet.getMergeCells();
            assertEquals(list.size(), 4);
            assertEquals(list.get(0), Dimension.of("B2:C2"));
            assertEquals(list.get(1), Dimension.of("E5:F8"));
            assertEquals(list.get(2), Dimension.of("A13:A20"));
            assertEquals(list.get(3), Dimension.of("B16:E17"));

            sheet = reader.sheet(1).asFullSheet();
            list = sheet.getMergeCells();
            assertEquals(list.size(), 2);
            assertEquals(list.get(0), Dimension.of("BM2:BQ11"));
            assertEquals(list.get(1), Dimension.of("A1:B26"));

            sheet = reader.sheet(2).asFullSheet();
            list = sheet.getMergeCells();
            assertEquals(list.size(), 2);
            assertEquals(list.get(0), Dimension.of("A16428:D16437"));
            assertEquals(list.get(1), Dimension.of("A1:K3"));

            sheet = reader.sheet(3).asFullSheet();
            list = sheet.getMergeCells();
            assertEquals(list.size(), 1);
            assertEquals(list.get(0), Dimension.of("A1:CF1434"));
        }
    }

    @Test public void testLargeMerge() throws IOException {
        try (ExcelReader reader = ExcelReader.read(testResourceRoot().resolve("largeMerged.xlsx"))) {
            FullSheet sheet = reader.sheet(0).asFullSheet();
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
        final String fileName = "Force Import.xlsx";
        Map<String, Object> data1 = new HashMap<>();
        data1.put("id", 1);
        data1.put("name", "abc");

        Map<String, Object> data2 = new HashMap<>();
        data2.put("id", 2);
        data2.put("name", "xyz");
        new Workbook()
            .addSheet(new ListMapSheet<>().setData(Arrays.asList(data1, data2)))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            List<U> list = reader.sheet(0).forceImport().dataRows().map(row -> row.to(U.class)).collect(Collectors.toList());
            assertEquals(list.size(), 2);
            assertEquals("1: abc", list.get(0).toString());
            assertEquals("2: xyz", list.get(1).toString());
        }
    }

    @Test public void testUpperCaseRead() throws IOException {
        final String fileName = "Upper case Reader test.xlsx";
        Map<String, Object> data1 = new HashMap<>();
        data1.put("ID", 1);
        data1.put("NAME", "abc");

        Map<String, Object> data2 = new HashMap<>();
        data2.put("ID", 2);
        data2.put("NAME", "xyz");

        new Workbook()
            .addSheet(new ListMapSheet<>(Arrays.asList(data1, data2)))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
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
        final String fileName = "Underline case Reader test.xlsx";
        Map<String, Object> data1 = new HashMap<>();
        data1.put("USER_ID", 1);
        data1.put("USER_NAME", "abc");

        Map<String, Object> data2 = new HashMap<>();
        data2.put("USER_ID", 2);
        data2.put("USER_NAME", "xyz");

        new Workbook()
            .addSheet(new ListMapSheet<>(Arrays.asList(data1, data2)))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
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
        final String fileName = "empty.xlsx";
        new Workbook().addSheet(new ListSheet<>()).writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            Sheet sheet = reader.sheet(0);
            List<U> list = sheet.header(1, 2).bind(U.class).rows().map(row -> (U) row.get()).collect(Collectors.toList());
            assertTrue(list.isEmpty());

            list = sheet.reset().header(1).rows().map(row -> row.to(U.class)).collect(Collectors.toList());
            assertTrue(list.isEmpty());

            list = sheet.reset().dataRows().map(row -> row.to(U.class)).collect(Collectors.toList());
            assertTrue(list.isEmpty());
        }
    }

    @Test public void testEntryMissKey() throws IOException {
        final String fileName = "test entry miss key.xlsx";
        List<ListObjectSheetTest.Item> expectList = ListObjectSheetTest.Item.randomTestData(10);
        new Workbook().addSheet(new ListSheet<ListObjectSheetTest.Item>(
                new Column("id"), new Column("name"))
                .setData(expectList))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
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
                            case DATE    : o = toDateTimeString(row.getDate(cell));          break;
                            default: o = row.getString(start);
                        }
                        if (StringUtil.isEmpty(e)) assertTrue(StringUtil.isEmpty(o));
                        else assertEquals(o, e);
                    }
                }
            }
        }
    }

    @Test public void testReadRelManager() throws IOException {
        try (ExcelReader reader = ExcelReader.read(testResourceRoot().resolve("template2.xlsx"))) {
            XMLSheet sheet = (XMLSheet) reader.sheet(2);
            RelManager rel = sheet.getRelManager();
            assertEquals(rel.size(), 4);
            Relationship rId1 = rel.getById("rId1");
            assertEquals(rId1.getTarget(), "../printerSettings/printerSettings3.bin");
            assertEquals(rId1.getType(), Const.Relationship.PRINTER_SETTINGS);

            Relationship rId2 = rel.getById("rId2");
            assertEquals(rId2.getTarget(), "../drawings/drawing1.xml");
            assertEquals(rId2.getType(), Const.Relationship.DRAWINGS);

            Relationship rId3 = rel.getById("rId3");
            assertEquals(rId3.getTarget(), "../drawings/vmlDrawing2.vml");
            assertEquals(rId3.getType(), Const.Relationship.VMLDRAWING);

            Relationship rId4 = rel.getById("rId4");
            assertEquals(rId4.getTarget(), "../comments2.xml");
            assertEquals(rId4.getType(), Const.Relationship.COMMENTS);
        }
    }

    @Test public void testReadComments() throws IOException {
        try (ExcelReader reader = ExcelReader.read(testResourceRoot().resolve("template2.xlsx"))) {
            Map<Long, Comment> commentMap = reader.sheet(2).asFullSheet().getComments();
            assertEquals(commentMap.size(), 31);
            Comment c1 = commentMap.get(coordinateToLong("C1"));
            assertNotNull(c1);
            assertEquals(c1.title, "Administrator:");
            assertEquals(c1.value, "如果有一票多个FBA号码一起发货，多出的FBA号请填写在C12备注栏！\n" +
                "\n" +
                "\n" +
                "如FBA123456788\n" +
                "（注意，不要U）");

            Comment j17 = commentMap.get(coordinateToLong("J17"));
            assertNotNull(j17);
            assertEquals(j17.value, "\n" +
                "填货物的单位CTN，即货物的箱数\n" +
                "\n" +
                "有多少箱就填多少箱，如果是混装产品对应箱数都写1\n" +
                "\n" +
                "不得合并单元格！\n");
        }
    }

    @Test public void testSpecifyHeaderRow() throws IOException {
        try (ExcelReader reader = ExcelReader.read(testResourceRoot().resolve("1.xlsx"))) {
            Cell[] cells = new Cell[6];
            cells[0] = new Cell(1).setString("channelId");
            cells[1] = new Cell(2).setString("game");
            cells[2] = new Cell();
            cells[3] = new Cell(4).setString("registrationTime");
            cells[4] = new Cell(5).setString("up30");
            cells[5] = new Cell(6).setString("vip");
            org.ttzero.excel.reader.Row headerRow = new org.ttzero.excel.reader.Row() {};
            headerRow.setCells(cells);
            List<A> list = reader.sheet(0).forceImport().header(1).bind(A.class, headerRow)
                .dataRows().map(row -> (A) row.get()).collect(Collectors.toList());
            assertEquals(list.size(), 94);
            assertEquals(list.get(0).toString(), "4\t极品飞车\t2018-11-21\ttrue\tF");
            assertEquals(list.get(10).toString(), "3\t守望先锋\t2018-11-21\ttrue\tI");
            assertEquals(list.get(20).toString(), "9\t守望先锋\t2018-11-21\ttrue\tP");
            assertEquals(list.get(30).toString(), "9\t怪物世界\t2018-11-21\ttrue\tH");
            assertEquals(list.get(40).toString(), "10\tLOL\t2018-11-21\ttrue\tB");
            assertEquals(list.get(50).toString(), "8\t怪物世界\t2018-11-21\ttrue\tK");
            assertEquals(list.get(60).toString(), "5\t怪物世界\t2018-11-21\ttrue\tP");
            assertEquals(list.get(70).toString(), "7\t怪物世界\t2018-11-21\tfalse\tA");
            assertEquals(list.get(80).toString(), "7\tWOW\t2018-11-21\ttrue\tX");
            assertEquals(list.get(90).toString(), "3\t极品飞车\t2018-11-21\ttrue\tP");
        }
    }

    @Test public void testCalcSheet() throws IOException {
        try (ExcelReader reader = ExcelReader.read(testResourceRoot().resolve("formula.xlsx"))) {
            reader.sheets().map(Sheet::asCalcSheet).forEach(sheet -> {
                Map<Long, String> formulasMap = FormulasLoader.load(testResourceRoot().resolve("expect/formula$" + sheet.getName() + "$formulas.txt"));
                Iterator<Row> it = sheet.iterator();
                while (it.hasNext()) {
                    Row row = it.next();

                    for (int start = row.getFirstColumnIndex(), end = row.getLastColumnIndex(); start < end; start++) {
                        String formula = formulasMap.get(((long) row.getRowNum()) << 16 | (start + 1));
                        assertTrue(formula == null || formula.equals(row.getFormula(start)));
                    }
                }
            });
        }
    }

    @Test public void testMergeSheet() throws IOException {
        try (ExcelReader reader = ExcelReader.read(testResourceRoot().resolve("fracture merged.xlsx"))) {
            MergeSheet sheet = reader.sheet(0).asMergeSheet();
            List<Dimension> mergeCells = sheet.getMergeCells();
            assertEquals(mergeCells.size(), 17);
            assertTrue(mergeCells.contains(Dimension.of("A1:A2")));
            assertTrue(mergeCells.contains(Dimension.of("B1:B2")));
            assertTrue(mergeCells.contains(Dimension.of("C1:C2")));
            assertTrue(mergeCells.contains(Dimension.of("D1:D2")));
            assertTrue(mergeCells.contains(Dimension.of("E1:E2")));
            assertTrue(mergeCells.contains(Dimension.of("F1:F2")));
            assertTrue(mergeCells.contains(Dimension.of("G1:G2")));
            assertTrue(mergeCells.contains(Dimension.of("H1:I1")));
            assertTrue(mergeCells.contains(Dimension.of("J1:K1")));
            assertTrue(mergeCells.contains(Dimension.of("L1:M1")));
            assertTrue(mergeCells.contains(Dimension.of("N1:O1")));
            assertTrue(mergeCells.contains(Dimension.of("P1:Q1")));
            assertTrue(mergeCells.contains(Dimension.of("R1:S1")));
            assertTrue(mergeCells.contains(Dimension.of("T1:U1")));
            assertTrue(mergeCells.contains(Dimension.of("V1:W1")));
            assertTrue(mergeCells.contains(Dimension.of("X1:Y1")));
            assertTrue(mergeCells.contains(Dimension.of("Z1:AA1")));
            Iterator<Row> iter = sheet.iterator();
            assertTrue(iter.hasNext());
            Row row0 = iter.next();
            assertEquals(row0.toString(), "姓名 | 二级机构名称 | 三级机构名称 | 四级机构名称 | 参与次数 | 日均参与率(%) | 日均得分 | 2021-07-01 | 2021-07-01 | 2021-07-02 | 2021-07-02 | 2021-07-03 | 2021-07-03 | 2021-07-04 | 2021-07-04 | 2021-07-05 | 2021-07-05 | 2021-07-06 | 2021-07-06 | 2021-07-07 | 2021-07-07 | 2021-07-08 | 2021-07-08 | 2021-07-09 | 2021-07-09 | 2021-07-10 | 2021-07-10");
            assertTrue(iter.hasNext());
            Row row1 = iter.next();
            assertEquals(row1.toString(), "姓名 | 二级机构名称 | 三级机构名称 | 四级机构名称 | 参与次数 | 日均参与率(%) | 日均得分 | 得分 | 考试时长 | 得分 | 考试时长 | 得分 | 考试时长 | 得分 | 考试时长 | 得分 | 考试时长 | 得分 | 考试时长 | 得分 | 考试时长 | 得分 | 考试时长 | 得分 | 考试时长 | 得分 | 考试时长");
        }
    }

    @Test public void testReadHeader() throws IOException {
        try (ExcelReader reader = ExcelReader.read(testResourceRoot().resolve("1.xlsx"))) {
            Sheet sheet = reader.sheet("header rows test");
            Iterator<Row> iter = sheet.header(1).iterator();
            Row headerRow = sheet.getHeader();
            assertTrue(headerRow.isBlank());
            assertTrue(iter.hasNext());
            assertEquals(iter.next().toString(), "渠道ID | 游戏 | 账号 | 注册时间 | 是否满30级 | VIP");
            assertTrue(iter.hasNext());
            assertEquals(iter.next().toString(), "4 | 极品飞车 | XuSu2gFg32 | 2018-11-21 | true | F");
            assertTrue(iter.hasNext());
            assertEquals(iter.next().toString(), "8 | 怪物世界 | kxwWgaB | 2018-11-21 | true | N");

            sheet.reset();
            iter = sheet.header(1, 2).iterator();
            headerRow = sheet.getHeader();
            assertTrue(headerRow.isBlank());
            assertTrue(iter.hasNext());
            assertEquals(iter.next().toString(), "渠道ID | 游戏 | 账号 | 注册时间 | 是否满30级 | VIP");
            assertTrue(iter.hasNext());
            assertEquals(iter.next().toString(), "4 | 极品飞车 | XuSu2gFg32 | 2018-11-21 | true | F");
            assertTrue(iter.hasNext());
            assertEquals(iter.next().toString(), "8 | 怪物世界 | kxwWgaB | 2018-11-21 | true | N");

            sheet.reset();
            List<Map<String, Object>> list = sheet.header(3).rows().map(Row::toMap).collect(Collectors.toList());
            assertEquals(2, list.size());
            Map<String, Object> row1 = list.get(0);
            assertEquals(row1.get("渠道ID"), 4);
            assertEquals(row1.get("游戏"), "极品飞车");
            assertEquals(row1.get("账号"), "XuSu2gFg32");
            assertEquals(row1.get("注册时间"), Timestamp.valueOf(LocalDate.of(2018, 11, 21).atStartOfDay()));
            assertEquals(row1.get("是否满30级"), true);
            assertEquals(row1.get("VIP"), "F");

            Map<String, Object> row2 = list.get(1);
            assertEquals(row2.get("渠道ID"), 8);
            assertEquals(row2.get("游戏"), "怪物世界");
            assertEquals(row2.get("账号"), "kxwWgaB");
            assertEquals(row2.get("注册时间"), Timestamp.valueOf(LocalDate.of(2018, 11, 21).atStartOfDay()));
            assertEquals(row2.get("是否满30级"), true);
            assertEquals(row2.get("VIP"), "N");

            sheet.reset();
            list = sheet.header(3, 4).rows().map(Row::toMap).collect(Collectors.toList());
            assertEquals(1, list.size());
            Map<String, Object> row3 = list.get(0);
            assertEquals(row3.get("渠道ID:4"), 8);
            assertEquals(row3.get("游戏:极品飞车"), "怪物世界");
            assertEquals(row3.get("账号:XuSu2gFg32"), "kxwWgaB");
            assertEquals(row3.get("注册时间:43425"), Timestamp.valueOf(LocalDate.of(2018, 11, 21).atStartOfDay()));
            assertEquals(row3.get("是否满30级:true"), true);
            assertEquals(row3.get("VIP:F"), "N");
        }
    }

    @Test public void testLastColumnIndex() throws IOException {
        try (ExcelReader reader = ExcelReader.read(testResourceRoot().resolve("#175.xlsx"))) {
            int[] expect = {0, 1, 0, 2, 0, 1, 0, 6, 6, 6};
            for (Iterator<Row> iter = reader.sheet(0).iterator(); iter.hasNext(); ) {
                Row row = iter.next();
                assertEquals(expect[row.getRowNum()], row.getLastColumnIndex());
            }
        }

        try (ExcelReader reader = ExcelReader.read(testResourceRoot().resolve("#354.xlsx"))) {
            int[] expect = {0, 27, 24, 5};
            for (Iterator<Row> iter = reader.sheet(0).iterator(); iter.hasNext(); ) {
                Row row = iter.next();
                assertEquals(expect[row.getRowNum()], row.getLastColumnIndex());
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

    public static class A {
        private int channelId;
        private String game;
        private LocalDate registrationTime;
        private boolean up30;
        private char vip;
        @Override
        public String toString() {
            return channelId + "\t" + game + "\t" + registrationTime + "\t" + up30 + "\t" + vip;
        }
    }
}

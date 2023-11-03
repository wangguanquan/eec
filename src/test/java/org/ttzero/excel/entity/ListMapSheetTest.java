/*
 * Copyright (c) 2017-2019, guanquan.wang@yandex.com All Rights Reserved.
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
import org.ttzero.excel.entity.style.Horizontals;
import org.ttzero.excel.entity.style.NumFmt;
import org.ttzero.excel.entity.style.PatternType;
import org.ttzero.excel.entity.style.Styles;
import org.ttzero.excel.reader.Drawings;
import org.ttzero.excel.reader.ExcelReader;
import org.ttzero.excel.reader.HeaderRow;
import org.ttzero.excel.reader.Row;
import org.ttzero.excel.util.StringUtil;

import java.awt.Color;
import java.io.IOException;
import java.math.BigDecimal;
import java.sql.Time;
import java.sql.Timestamp;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.time.ZoneOffset;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.UUID;
import java.util.stream.Collectors;

/**
 * @author guanquan.wang at 2019-04-28 19:16
 */
public class ListMapSheetTest extends WorkbookTest {

    @Test public void testWrite() throws IOException {
        String fileName = "test map.xlsx";
        List<Map<String, ?>> expectList = createTestData();
        new Workbook()
            .addSheet(new ListMapSheet(expectList))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            List<Map<String, ?>> list = reader.sheet(0).dataRows().map(Row::toMap).collect(Collectors.toList());
            assert expectList.size() == list.size();
            for (int i = 0, len = expectList.size(); i < len; i++) {
                Map<String, ?> expect = expectList.get(i), e = list.get(i);
                assert expect.equals(e);
            }
        }
    }

    @Test public void testAllType() throws IOException {
        String fileName = "test all type map.xlsx";
        List<Map<String, ?>> expectList = createAllTypeData();
        new Workbook()
            .setAutoSize(true)
            .addSheet(new ListMapSheet(expectList))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            List<Map<String, ?>> list = reader.sheet(0).dataRows().map(Row::toMap).collect(Collectors.toList());
            assertAllTypes(expectList, list);
        }
    }

    @Test public void testStyleDesign4Map() throws IOException {
        String fileName = "Map标识行样式.xlsx";
        List<Map<String, ?>> expectList = createAllTypeData(100);
        new Workbook()
                .addSheet(new ListMapSheet("Map", expectList).setStyleProcessor((map, style, sst) -> {
                    if ((Boolean) map.get("bv")) {
                        style = sst.modifyFill(style, new Fill(PatternType.solid, Color.green));
                    }
                    return style;
                }))
                .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            org.ttzero.excel.reader.Sheet sheet = reader.sheet(0).header(1);
            assert "Map".equals(sheet.getName());
            Iterator<org.ttzero.excel.reader.Row> iter = sheet.iterator();
            for (Map<String, ?> expect : expectList) {
                assert iter.hasNext();
                org.ttzero.excel.reader.Row row = iter.next();
                Map<String, ?> e = row.toMap();
                assertAllType(expect, e);

                boolean bv = (Boolean) expect.get("bv");
                Styles styles = row.getStyles();
                for (int i = row.getFirstColumnIndex(), j = row.getLastColumnIndex(); i < j; i++) {
                    int styleIndex = row.getCellStyle(i);
                    Fill fill = styles.getFill(styleIndex);
                    if (bv) {
                        assert fill != null && fill.getPatternType() == PatternType.solid && fill.getFgColor().equals(Color.green);
                    } else {
                        assert fill == null || fill.getPatternType() == PatternType.none;
                    }
                }
            }
        }
    }

    @Test public void testStyleDesign4Map2() throws IOException {
        String fileName = "Map标识行样式2.xlsx";
        List<Map<String, ?>> expectList = createAllTypeData(100);
        new Workbook()
            .addSheet(new ListMapSheet("Map", expectList
                , new Column("boolean", "bv", boolean.class)
                , new Column("char", "cv", char.class)
                , new Column("short", "sv", short.class)
                , new Column("int", "nv", int.class).setStyleProcessor((n,s,sst) -> ((int) n) < 0 ? sst.modifyHorizontal(s, Horizontals.LEFT) : s).setNumFmt("¥0.00_);[Red]-¥0.00_);¥0_)")
                , new Column("long", "lv", long.class)
                , new Column("LocalDateTime", "ldtv", LocalDateTime.class)
                , new Column("LocalTime", "ltv", LocalTime.class)).setStyleProcessor((map, style, sst)->{
                if ((Boolean) map.get("bv")) {
                    style = sst.modifyFill(style, new Fill(PatternType.solid, Color.green));
                }
                return style;
            }))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            org.ttzero.excel.reader.Sheet sheet = reader.sheet(0).header(1);
            assert "Map".equals(sheet.getName());
            org.ttzero.excel.reader.HeaderRow header = (HeaderRow) sheet.getHeader();
            assert "boolean".equals(header.get(0));
            assert "char".equals(header.get(1));
            assert "short".equals(header.get(2));
            assert "int".equals(header.get(3));
            assert "long".equals(header.get(4));
            assert "LocalDateTime".equals(header.get(5));
            assert "LocalTime".equals(header.get(6));
            Iterator<org.ttzero.excel.reader.Row> iter = sheet.iterator();
            for (Map<String, ?> expect : expectList) {
                assert iter.hasNext();
                org.ttzero.excel.reader.Row row = iter.next();
                Map<String, ?> e = row.toMap();

                assert expect.get("bv").equals(e.get("boolean"));
                assert expect.get("cv").toString().equals(e.get("char").toString());
                assert expect.get("sv").toString().equals(e.get("short").toString());
                assert expect.get("nv").toString().equals(e.get("int").toString());
                assert expect.get("lv").toString().equals(e.get("long").toString());
                LocalDateTime ldtv1 = (LocalDateTime) expect.get("ldtv");
                Timestamp ldtv2 = (Timestamp) e.get("LocalDateTime");
                assert Timestamp.valueOf(ldtv1).getTime() / 1000 == ldtv2.getTime() / 1000;
                LocalTime ltv1 = (LocalTime) expect.get("ltv");
                Time ltv2 = (Time) e.get("LocalTime");
                assert String.valueOf(Time.valueOf(ltv1)).equals(String.valueOf(ltv2));

                boolean bv = (Boolean) expect.get("bv");
                Styles styles = row.getStyles();
                for (int i = row.getFirstColumnIndex(), j = row.getLastColumnIndex(); i < j; i++) {
                    int styleIndex = row.getCellStyle(i);
                    Fill fill = styles.getFill(styleIndex);
                    if (bv) {
                        assert fill != null && fill.getPatternType() == PatternType.solid && fill.getFgColor().equals(Color.green);
                    } else {
                        assert fill == null || fill.getPatternType() == PatternType.none;
                    }
                }

                int styleIndex3 = row.getCellStyle(3), horizontals = styles.getHorizontal(styleIndex3) << Styles.INDEX_HORIZONTAL;
                NumFmt numFmt = styles.getNumFmt(styleIndex3);
                assert "¥0.00_);[Red]\\-¥0.00_);¥0_)".equals(numFmt.getCode());
                if ((Integer) expect.get("nv") < 0) {
                    assert Horizontals.LEFT == horizontals;
                } else {
                    assert Horizontals.RIGHT == horizontals;
                }
            }
        }
    }

    @Test public void testHeaderColumn() throws IOException {
        String fileName = "test header column map.xlsx";
        List<Map<String, ?>> expectList = createAllTypeData();
        new Workbook()
            .addSheet(new ListMapSheet(expectList
                , new Column("boolean", "bv", boolean.class)
                , new Column("char", "cv", char.class)
                , new Column("short", "sv", short.class)
                , new Column("int", "nv", int.class)
                , new Column("long", "lv", long.class)
                , new Column("float", "fv", float.class)
                , new Column("double", "dv", double.class)
                , new Column("string", "s", String.class)
                , new Column("decimal", "mv", BigDecimal.class)
                , new Column("date", "av", Date.class)
                , new Column("timestamp", "iv", Timestamp.class)
                , new Column("time", "tv", Time.class)
                , new Column("LocalDate", "ldv", LocalDate.class)
                , new Column("LocalDateTime", "ldtv", LocalDateTime.class)
                , new Column("LocalTime", "ltv", LocalTime.class)
            ))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            List<Map<String, ?>> list = reader.sheet(0).dataRows().map(Row::toMap).collect(Collectors.toList());
            assert expectList.size() == list.size();
            for (int i = 0, len = expectList.size(); i < len; i++) {
                Map<String, ?> expect = expectList.get(i), e = list.get(i);
                assertAllTypeFullKey(expect, e);
            }
        }
    }

    @Test public void testHeaderColumnBox() throws IOException {
        String fileName = "test header column box type map.xlsx";
        List<Map<String, ?>> expectList = createAllTypeData();
        new Workbook()
            .addSheet(new ListMapSheet(expectList
                , new Column("Character", "cv", Character.class)
                , new Column("Short", "sv", Short.class)
                , new Column("Integer", "nv", Integer.class)
                , new Column("Long", "lv", Long.class)
                , new Column("Float", "fv", Float.class)
                , new Column("Double", "dv", Double.class)
            ))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            org.ttzero.excel.reader.Sheet sheet = reader.sheet(0).header(1);
            org.ttzero.excel.reader.HeaderRow header = (HeaderRow) sheet.getHeader();
            assert "Character".equals(header.get(0));
            assert "Short".equals(header.get(1));
            assert "Integer".equals(header.get(2));
            assert "Long".equals(header.get(3));
            assert "Float".equals(header.get(4));
            assert "Double".equals(header.get(5));
            Iterator<org.ttzero.excel.reader.Row> iter = sheet.iterator();
            for (Map<String, ?> expect : expectList) {
                assert iter.hasNext();
                org.ttzero.excel.reader.Row row = iter.next();
                Map<String, ?> e = row.toMap();

                assert expect.get("cv").toString().equals(e.get("Character").toString());
                assert expect.get("sv").toString().equals(e.get("Short").toString());
                assert expect.get("nv").toString().equals(e.get("Integer").toString());
                assert expect.get("lv").toString().equals(e.get("Long").toString());
                assert Float.compare((Float) expect.get("fv"), Float.parseFloat(e.get("Float").toString())) == 0;
                assert Double.compare((Double) expect.get("dv"), Double.parseDouble(e.get("Double").toString())) == 0;
            }
        }
    }

    @Test public void testConstructor1() throws IOException {
        String fileName = "test list map sheet Constructor1.xlsx";
        new Workbook()
            .setAutoSize(true)
            .addSheet(new ListMapSheet())
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            assert reader.sheet(0).rows().count() == 0L;
        }
    }

    @Test public void testConstructor2() throws IOException {
        String fileName = "test list map sheet Constructor2.xlsx";
        List<Map<String, ?>> expectList = createTestData(10);
        new Workbook()
            .setAutoSize(true)
            .addSheet(new ListMapSheet("Map").setData(expectList))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            List<Map<String, ?>> list = reader.sheet(0).dataRows().map(Row::toMap).collect(Collectors.toList());
            assert expectList.size() == list.size();
            for (int i = 0, len = expectList.size(); i < len; i++) {
                Map<String, ?> expect = expectList.get(i), e = list.get(i);
                assert expect.equals(e);
            }
        }
    }

    @Test public void testConstructor3() throws IOException {
        String fileName = "test list map sheet Constructor3.xlsx";
        List<Map<String, ?>> expectList = createAllTypeData(10);
        new Workbook()
            .setAutoSize(true)
            .addSheet(new ListMapSheet("Map"
                , new Column("boolean", "bv", boolean.class)
                , new Column("char", "cv", char.class)
                , new Column("short", "sv", short.class)
                , new Column("int", "nv", int.class)
                , new Column("long", "lv", long.class)
                , new Column("float", "fv", float.class)
                , new Column("double", "dv", double.class)
                , new Column("string", "s", String.class)
                , new Column("decimal", "mv", BigDecimal.class)
                , new Column("date", "av", Date.class)
                , new Column("timestamp", "iv", Timestamp.class)
                , new Column("time", "tv", Time.class)
                , new Column("LocalDate", "ldv", LocalDate.class)
                , new Column("LocalDateTime", "ldtv", LocalDateTime.class)
                , new Column("LocalTime", "ltv", LocalTime.class)
            ).setData(expectList))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            List<Map<String, ?>> list = reader.sheet(0).dataRows().map(Row::toMap).collect(Collectors.toList());
            assert expectList.size() == list.size();
            for (int i = 0, len = expectList.size(); i < len; i++) {
                Map<String, ?> expect = expectList.get(i), e = list.get(i);
                assertAllTypeFullKey(expect, e);
            }
        }
    }

    @Test public void testConstructor4() throws IOException {
        String fileName = "test list map sheet Constructor4.xlsx";
        List<Map<String, ?>> expectList = createAllTypeData(10);
        new Workbook()
            .setAutoSize(true)
            .addSheet(new ListMapSheet("Map", WaterMark.of(author)
                , new Column("boolean", "bv", boolean.class)
                , new Column("char", "cv", char.class)
                , new Column("short", "sv", short.class)
                , new Column("int", "nv", int.class)
                , new Column("long", "lv", long.class)
                , new Column("float", "fv", float.class)
                , new Column("double", "dv", double.class)
                , new Column("string", "s", String.class)
                , new Column("decimal", "mv", BigDecimal.class)
                , new Column("date", "av", Date.class)
                , new Column("timestamp", "iv", Timestamp.class)
                , new Column("time", "tv", Time.class)
                , new Column("LocalDate", "ldv", LocalDate.class)
                , new Column("LocalDateTime", "ldtv", LocalDateTime.class)
                , new Column("LocalTime", "ltv", LocalTime.class))
                .setData(expectList))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            org.ttzero.excel.reader.Sheet sheet = reader.sheet(0);
            // Water Mark
            List<Drawings.Picture> pictures = sheet.listPictures();
            assert pictures.size() == 1;
            assert pictures.get(0).isBackground();

            List<Map<String, ?>> list = sheet.dataRows().map(Row::toMap).collect(Collectors.toList());
            assert expectList.size() == list.size();
            for (int i = 0, len = expectList.size(); i < len; i++) {
                Map<String, ?> expect = expectList.get(i), e = list.get(i);
                assertAllTypeFullKey(expect, e);
            }
        }
    }

    @Test public void testConstructor5() throws IOException {
        String fileName = "test list map sheet Constructor5.xlsx";
        List<Map<String, ?>> expectList = createAllTypeData(10);
        new Workbook()
            .setAutoSize(true)
            .addSheet(new ListMapSheet(expectList))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            org.ttzero.excel.reader.Sheet sheet = reader.sheet(0);
            List<Map<String, ?>> list = sheet.dataRows().map(Row::toMap).collect(Collectors.toList());
            assert expectList.size() == list.size();
            for (int i = 0, len = expectList.size(); i < len; i++) {
                Map<String, ?> expect = expectList.get(i), e = list.get(i);
                assertAllType(expect, e);
            }
        }
    }

    @Test public void testConstructor6() throws IOException {
        String fileName = "test list map sheet Constructor6.xlsx";
        List<Map<String, ?>> expectList = createAllTypeData(10);
        new Workbook()
            .setAutoSize(true)
            .addSheet(new ListMapSheet("Map", expectList))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            org.ttzero.excel.reader.Sheet sheet = reader.sheet(0);
            assert "Map".equals(sheet.getName());
            List<Map<String, ?>> list = sheet.dataRows().map(Row::toMap).collect(Collectors.toList());
            assert expectList.size() == list.size();
            for (int i = 0, len = expectList.size(); i < len; i++) {
                Map<String, ?> expect = expectList.get(i), e = list.get(i);
                assertAllType(expect, e);
            }
        }
    }

    @Test public void testConstructor8() throws IOException {
        String fileName = "test list map sheet Constructor8.xlsx";
        List<Map<String, ?>> expectList = createTestData(10);
        new Workbook()
            .setAutoSize(true)
            .addSheet(new ListMapSheet("MAP", expectList
                , new Column("ID", "id", int.class)
                , new Column("NAME", "name", String.class)))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            org.ttzero.excel.reader.Sheet sheet = reader.sheet(0);
            assert "MAP".equals(sheet.getName());
            List<Map<String, ?>> list = sheet.dataRows().map(Row::toMap).collect(Collectors.toList());
            assert expectList.size() == list.size();
            for (int i = 0, len = expectList.size(); i < len; i++) {
                Map<String, ?> expect = expectList.get(i), e = list.get(i);
                assert expect.get("id").equals(e.get("ID"));
                assert expect.get("name").equals(e.get("NAME"));
            }
        }
    }

    @Test public void testConstructor9() throws IOException {
        String fileName = "test list map sheet Constructor9.xlsx";
        List<Map<String, ?>> expectList = createTestData(10);
        new Workbook()
            .setAutoSize(true)
            .addSheet(new ListMapSheet(expectList
                , WaterMark.of(author)
                , new Column("ID", "id")
                , new Column("NAME", "name")))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            org.ttzero.excel.reader.Sheet sheet = reader.sheet(0);

            // Water Mark
            List<Drawings.Picture> pictures = sheet.listPictures();
            assert pictures.size() == 1;
            assert pictures.get(0).isBackground();

            List<Map<String, ?>> list = sheet.dataRows().map(Row::toMap).collect(Collectors.toList());
            assert expectList.size() == list.size();
            for (int i = 0, len = expectList.size(); i < len; i++) {
                Map<String, ?> expect = expectList.get(i), e = list.get(i);
                assert expect.get("id").equals(e.get("ID"));
                assert expect.get("name").equals(e.get("NAME"));
            }
        }
    }

    @Test public void testConstructor10() throws IOException {
        String fileName = "test list map sheet Constructor10.xlsx";
        List<Map<String, ?>> expectList = createTestData(10);
        new Workbook()
            .setAutoSize(true)
            .addSheet(new ListMapSheet("MAP"
                , expectList
                , WaterMark.of(author)
                , new Column("ID", "id", int.class)
                , new Column("NAME", "name", String.class)))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            org.ttzero.excel.reader.Sheet sheet = reader.sheet(0);
            assert "MAP".equals(sheet.getName());

            // Water Mark
            List<Drawings.Picture> pictures = sheet.listPictures();
            assert pictures.size() == 1;
            assert pictures.get(0).isBackground();

            List<Map<String, ?>> list = sheet.dataRows().map(Row::toMap).collect(Collectors.toList());
            assert expectList.size() == list.size();
            for (int i = 0, len = expectList.size(); i < len; i++) {
                Map<String, ?> expect = expectList.get(i), e = list.get(i);
                assert expect.get("id").equals(e.get("ID"));
                assert expect.get("name").equals(e.get("NAME"));
            }
        }
    }

    @Test public void testArray() throws IOException {
        String fileName = "ListMapSheet Array Map.xlsx";
        List<Map<String, ?>> expectList;
        Map<String, Object> data1 = new HashMap<>();
        data1.put("id", 1);
        data1.put("name", "abc");

        Map<String, Object> data2 = new HashMap<>();
        data2.put("id", 2);
        data2.put("name", "xyz");
        new Workbook()
            .addSheet(new ListMapSheet().setData(expectList = Arrays.asList(data1, data2)))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            org.ttzero.excel.reader.Sheet sheet = reader.sheet(0);
            List<Map<String, ?>> list = sheet.dataRows().map(Row::toMap).collect(Collectors.toList());
            assert expectList.size() == list.size();
            for (int i = 0, len = expectList.size(); i < len; i++) {
                Map<String, ?> expect = expectList.get(i), e = list.get(i);
                assert expect.equals(e);
            }
        }
    }

    @Test public void testSingleList() throws IOException {
        String fileName = "ListMapSheet Single List Map.xlsx";
        List<Map<String, ?>> expectList;
        Map<String, Object> data = new HashMap<>();
        data.put("id", 1);
        data.put("name", "abc");

        new Workbook()
            .addSheet(new ListMapSheet().setData(expectList = Collections.singletonList(data)))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            org.ttzero.excel.reader.Sheet sheet = reader.sheet(0);
            List<Map<String, ?>> list = sheet.dataRows().map(Row::toMap).collect(Collectors.toList());
            assert expectList.size() == list.size();
            for (int i = 0, len = expectList.size(); i < len; i++) {
                Map<String, ?> expect = expectList.get(i), e = list.get(i);
                assert expect.equals(e);
            }
        }
    }

    @Test public void testNullValue() throws IOException {
        String fileName = "test map null value.xlsx";
        List<Map<String, ?>> expectList = createNullTestData(10);
        new Workbook()
            .addSheet(new ListMapSheet(expectList))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            org.ttzero.excel.reader.Sheet sheet = reader.sheet(0);
            List<Map<String, ?>> list = sheet.dataRows().map(Row::toMap).collect(Collectors.toList());
            assert expectList.size() == list.size();
            for (int i = 0, len = expectList.size(); i < len; i++) {
                Map<String, ?> expect = expectList.get(i), e = list.get(i);
                assert expect.get("id").equals(e.get("id"));
                assert e.get("name") == null || StringUtil.isEmpty(e.get("name").toString());
            }
        }
    }

    // Issue #93
    @Test public void testListMapSheet_93() throws IOException {
        String fileName = "Issue#93 List Map.xlsx";
        List<Map<String, ?>> expectList = new ArrayList<>();
        new Workbook().addSheet(new ListMapSheet() {
            private int i;
            @Override
            protected List<Map<String, ?>> more() {
                List<Map<String, ?>> list = i++ < 10 ? createAllTypeData(30) : null;
                if (list != null) expectList.addAll(list);
                return list;
            }
        }).writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            List<Map<String, ?>> list = reader.sheet(0).dataRows().map(Row::toMap).collect(Collectors.toList());
            assertAllTypes(expectList, list);
        }
    }

    @Test public void test_161() throws IOException {
        String fileName = "Issue#161.xlsx";
        List<Map<String, ?>> expectList = new ArrayList<>();
        new Workbook().addSheet(new ListMapSheet() {
            private int i = 0;
            @Override
            protected List<Map<String, ?>> more() {
                // Only write one row
                if (i++ > 0) return null;
                List<Map<String, ?>> list = new ArrayList<>();
                Map<String, Object> map = new HashMap<>();
                map.put("uuid", UUID.randomUUID().toString());
                map.put("hobbies", new ArrayList<String>() {{
                    add("张");
                    add("李");
                }});
                map.put("sex", "男");
                final int len = 4095;
                char[] chars = new char[len];
                Arrays.fill(chars, 'a');
                // java.nio.BufferOverflowException occur when the cell value length large than 2045
                map.put("name", new String(chars, 0, len));
                map.put("age", 24);
                map.put("createDate", new Date(1535444725000L).toInstant().atOffset(ZoneOffset.of("+8")).toLocalDateTime());

                list.add(map);
                expectList.addAll(list);
                return list;
            }
        }).writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            List<Map<String, ?>> list = reader.sheet(0).dataRows().map(Row::toMap).collect(Collectors.toList());
            assert list.size() == expectList.size();
            Map<String, ?> expect = expectList.get(0), e = list.get(0);
            assert expect.get("uuid").equals(e.get("uuid"));
            assert expect.get("hobbies").toString().equals(e.get("hobbies"));
            assert expect.get("sex").equals(e.get("sex"));
            assert expect.get("name").equals(e.get("name"));
            assert expect.get("age").equals(e.get("age"));
            LocalDateTime ldtv1 = (LocalDateTime) expect.get("createDate");
            Timestamp ldtv2 = (Timestamp) e.get("createDate");
            assert Timestamp.valueOf(ldtv1).getTime() / 1000 == ldtv2.getTime() / 1000;
        }
    }

    @Test public void testWrapText() throws IOException {
        String fileName = "MAP WRAP TEXT.xlsx";
        List<Map<String, ?>> expectList = createTestData(10);
        new Workbook()
                .addSheet(new ListMapSheet(expectList
                    , new Column("ID", "id", int.class)
                    , new Column("NAME", "name", String.class).setWrapText(true)
                ))
                .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            List<Map<String, ?>> list = reader.sheet(0).dataRows().map(Row::toMap).collect(Collectors.toList());
            assert expectList.size() == list.size();
            for (int i = 0, len = expectList.size(); i < len; i++) {
                Map<String, ?> expect = expectList.get(i), e = list.get(i);
                assert expect.get("id").equals(e.get("ID"));
                assert expect.get("name").equals(e.get("NAME"));
            }
        }
    }

    @Test(expected = TooManyColumnsException.class) public void testOverLargeOrderColumn() throws IOException {
        new Workbook("test list map sheet Constructor8", author)
                .setAutoSize(true)
                .addSheet(new ListMapSheet("MAP", createTestData(10)
                        , new Column("ID", "id", int.class).setColIndex(9999999)
                        , new Column("NAME", "name", String.class)))
                .writeTo(defaultTestPath);
    }

    @Test public void test257() throws IOException {
        String fileName = "Issue#257.xlsx";
        List<Map<String, ?>> expectList = new ArrayList<>();
        expectList.add(new HashMap<String, String>(){{put("sub1", "moban1");}});
        expectList.add(new HashMap<String, String>(){{put("sub2", "moban2");}});
        expectList.add(new HashMap<String, String>(){{put("sub3", "moban3");}});

        new Workbook().addSheet(new ListMapSheet(expectList
                , new Column("ID", "id")
                , new Column("子表单", "sub1")
                , new Column("模板2", "sub2")
                , new Column("模板3", "sub3")
                , new Column("abc", "abc")
                , new Column("模板2", "sub2")
                , new Column("xx", "xx")
                , new Column("xyz", "xyz")
        )).writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            org.ttzero.excel.reader.HeaderRow header = (HeaderRow) reader.sheet(0).header(1).getHeader();
            assert "ID".equals(header.get(0));
            assert "子表单".equals(header.get(1));
            assert "模板2".equals(header.get(2));
            assert "模板3".equals(header.get(3));
            assert "abc".equals(header.get(4));
            assert "模板2".equals(header.get(5));
            assert "xx".equals(header.get(6));
            assert "xyz".equals(header.get(7));
            List<Map<String, ?>> list = reader.sheet(0).dataRows().map(Row::toMap).collect(Collectors.toList());
            assert expectList.size() == list.size();
            Map<String, ?> expect = expectList.get(0), e = list.get(0);
            assert expect.get("sub1").equals(e.get("子表单"));
            expect = expectList.get(1); e = list.get(1);
            assert expect.get("sub2").equals(e.get("模板2"));
            expect = expectList.get(2); e = list.get(2);
            assert expect.get("sub3").equals(e.get("模板3"));
            assert e.get("ID") == null || StringUtil.isEmpty(e.get("ID").toString());
            assert e.get("abc") == null || StringUtil.isEmpty(e.get("abc").toString());
            assert e.get("xx") == null || StringUtil.isEmpty(e.get("xx").toString());
            assert e.get("xyz") == null || StringUtil.isEmpty(e.get("xyz").toString());
        }
    }

    @Test public void testNullInListMap() throws IOException {
        String fileName = "Null in list map.xlsx";
        List<Map<String, ?>> expectList = createTestData(10);
        expectList.add(0, null);
        expectList.add(3, null);
        expectList.add(null);
        new Workbook().addSheet(new ListMapSheet(expectList)).writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            Iterator<org.ttzero.excel.reader.Row> iter = reader.sheet(0).header(1).iterator();
            for (Map<String, ?> expect : expectList) {
                assert iter.hasNext();
                Row row = iter.next();
                if (expect == null || expect.isEmpty()) {
                    assert row.isBlank();
                } else {
                    assert expect.equals(row.toMap());
                }
            }
        }
    }

    @Test public void testLargeColumns() throws IOException {
        int len = 1436;
        List<Map<String, ?>> expectList = new ArrayList<>(len);
        for (int i = 0; i < len; i++) {
            Map<String, String> map = new LinkedHashMap<>();
            for (int j = 0; j < 500; j++) {
                map.put("key" + j, getRandomString());
            }
            expectList.add(map);
        }

        new Workbook().addSheet(new ListMapSheet(expectList)).writeTo(defaultTestPath.resolve("large map.xlsx"));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("large map.xlsx"))) {
            List<Map<String, ?>> list = reader.sheet(0).dataRows().map(Row::toMap).collect(Collectors.toList());
            assert expectList.size() == list.size();
            for (int i = 0; i < len; i++) {
                Map<String, ?> expect = expectList.get(i), e = list.get(i);
                assert expect.equals(e);
            }
        }
    }

    @Test public void testSpecifyRowWrite() throws IOException {
        List<Map<String, ?>> list = createTestData(10);
        new Workbook().setAutoSize(true)
            .addSheet(new ListMapSheet(list).setStartRowIndex(5))
            .writeTo(defaultTestPath.resolve("test specify row 5 ListMapSheet.xlsx"));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("test specify row 5 ListMapSheet.xlsx"))) {
            List<Map<String, ?>> readList = reader.sheet(0).header(5).rows().map(Row::toMap).collect(Collectors.toList());
            assert list.size() == readList.size();
            for (int i = 0, len = list.size(); i < len; i++) {
                Map<String, ?> r = readList.get(i), w = list.get(i);
                assert r.equals(w);
            }
        }
    }

    @Test public void testSpecifyRowStayA1Write() throws IOException {
        List<Map<String, ?>> list = createTestData(10);
        new Workbook().setAutoSize(true)
            .addSheet(new ListMapSheet(list).setStartRowIndex(5, false))
            .writeTo(defaultTestPath.resolve("test specify row 5 stay A1 ListMapSheet.xlsx"));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("test specify row 5 stay A1 ListMapSheet.xlsx"))) {
            List<Map<String, ?>> readList = reader.sheet(0).header(5).rows().map(Row::toMap).collect(Collectors.toList());
            assert list.size() == readList.size();
            for (int i = 0, len = list.size(); i < len; i++) {
                Map<String, ?> r = readList.get(i), w = list.get(i);
                assert r.equals(w);
            }
        }
    }

    @Test public void testSpecifyRowAndColWrite() throws IOException {
        List<Map<String, ?>> list = createTestData(10);
        new Workbook().setAutoSize(true)
            .addSheet(new ListMapSheet(list
                , new Column("id", "id").setColIndex(3)
                , new Column("name", "name").setColIndex(4))
                .setStartRowIndex(5))
            .writeTo(defaultTestPath.resolve("test specify row and col ListMapSheet.xlsx"));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("test specify row and col ListMapSheet.xlsx"))) {
            List<Map<String, ?>> readList = reader.sheet(0).header(5).rows().map(Row::toMap).collect(Collectors.toList());
            assert list.size() == readList.size();
            for (int i = 0, len = list.size(); i < len; i++) {
                Map<String, ?> r = readList.get(i), w = list.get(i);
                assert r.size() == w.size();
                assert r.get("id").equals(w.get("id"));
                assert r.get("name").equals(w.get("name"));
            }
        }
    }

    @Test public void testSpecifyRowAndColStayA1Write() throws IOException {
        List<Map<String, ?>> list = createTestData(10);
        new Workbook().setAutoSize(true)
            .addSheet(new ListMapSheet(list
                , new Column("id", "id").setColIndex(3)
                , new Column("name", "name").setColIndex(4))
                .setStartRowIndex(5, false))
            .writeTo(defaultTestPath.resolve("test specify row and cel stay A1 ListMapSheet.xlsx"));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("test specify row and cel stay A1 ListMapSheet.xlsx"))) {
            List<Map<String, ?>> readList = reader.sheet(0).header(5).rows().map(Row::toMap).collect(Collectors.toList());
            assert list.size() == readList.size();
            for (int i = 0, len = list.size(); i < len; i++) {
                Map<String, ?> r = readList.get(i), w = list.get(i);
                assert r.equals(w);
            }
        }
    }

    public static List<Map<String, ?>> createTestData() {
        int size = random.nextInt(100) + 1;
        return createTestData(size);
    }

    public static List<Map<String, ?>> createTestData(int size) {
        List<Map<String, ?>> list = new ArrayList<>(size);
        for (int i = 0; i < size; i++) {
            Map<String, Object> map = new LinkedHashMap<>();
            map.put("id", random.nextInt());
            map.put("name", getRandomString());
            list.add(map);
        }
        return list;
    }

    public static List<Map<String, ?>> createAllTypeData() {
        int size = random.nextInt(100) + 1;
        return createAllTypeData(size);
    }

    public static List<Map<String, ?>> createAllTypeData(int size) {
        List<Map<String, ?>> list = new ArrayList<>(size);
        for (int i = 0; i < size; i++) {
            Map<String, Object> map = new LinkedHashMap<>();
            map.put("bv", random.nextInt(10) == 6);
            map.put("cv", charArray[random.nextInt(charArray.length)]);
            map.put("sv", (short) (random.nextInt() & 0xFFFF));
            map.put("nv", random.nextInt());
            map.put("lv", random.nextLong());
            map.put("fv", random.nextFloat());
            map.put("dv", random.nextDouble());
            map.put("s", getRandomString());
            map.put("mv", BigDecimal.valueOf(random.nextDouble()));
            map.put("av", new Date());
            map.put("iv", new Timestamp(System.currentTimeMillis() - random.nextInt(9999999)));
            map.put("tv", new Time(random.nextLong()));
            map.put("ldv", LocalDate.now());
            map.put("ldtv", LocalDateTime.now());
            map.put("ltv", LocalTime.now());
            list.add(map);
        }
        return list;
    }

    public static List<Map<String, ?>> createNullTestData(int size) {
        List<Map<String, ?>> list = new ArrayList<>(size);
        for (int i = 0; i < size; i++) {
            Map<String, Object> map = new LinkedHashMap<>();
            map.put("id", random.nextInt());
            map.put("name", null);
            list.add(map);
        }
        return list;
    }

    static void assertAllTypes(List<Map<String, ?>> expectList, List<Map<String, ?>> list) {
        assert expectList.size() == list.size();
        for (int i = 0, len = expectList.size(); i < len; i++) {
            Map<String, ?> expect = expectList.get(i), e = list.get(i);
            assertAllType(expect, e);
        }
    }

    static void assertAllType(Map<String, ?> expect, Map<String, ?> e) {
        assert expect.size() == e.size();
        assert expect.get("bv").equals(e.get("bv"));
        assert expect.get("cv").toString().equals(e.get("cv").toString());
        assert expect.get("sv").toString().equals(e.get("sv").toString());
        assert expect.get("nv").toString().equals(e.get("nv").toString());
        assert expect.get("lv").toString().equals(e.get("lv").toString());
        assert Float.compare((Float) expect.get("fv"), Float.parseFloat(e.get("fv").toString())) == 0;
        assert Double.compare((Double) expect.get("dv"), Double.parseDouble(e.get("dv").toString())) == 0;
        assert expect.get("s").equals(e.get("s"));
        assert ((BigDecimal) expect.get("mv")).setScale(4, BigDecimal.ROUND_HALF_DOWN).equals(new BigDecimal(e.get("mv").toString()).setScale(4, BigDecimal.ROUND_HALF_DOWN));
        Date av1 = (Date) expect.get("av"), av2 = (Date) e.get("av");
        assert av1.getTime() / 1000 == av2.getTime() / 1000;
        Date iv1 = (Date) expect.get("iv"), iv2 = (Date) e.get("iv");
        assert iv1.getTime() / 1000 == iv2.getTime() / 1000;
        Time tv1 = (Time) expect.get("tv"), tv2 = (Time) e.get("tv");
        assert String.valueOf(tv1).equals(String.valueOf(tv2));
        LocalDate ldv1 = (LocalDate) expect.get("ldv");
        Timestamp ldv2 = (Timestamp) e.get("ldv");
        assert ldv1.equals(ldv2.toLocalDateTime().toLocalDate());
        LocalDateTime ldtv1 = (LocalDateTime) expect.get("ldtv");
        Timestamp ldtv2 = (Timestamp) e.get("ldtv");
        assert Timestamp.valueOf(ldtv1).getTime() / 1000 == ldtv2.getTime() / 1000;
        LocalTime ltv1 = (LocalTime) expect.get("ltv");
        Time ltv2 = (Time) e.get("ltv");
        assert String.valueOf(Time.valueOf(ltv1)).equals(String.valueOf(ltv2));
    }

    static void assertAllTypeFullKey(Map<String, ?> expect, Map<String, ?> e) {
        assert expect.size() == e.size();
        assert expect.get("bv").equals(e.get("boolean"));
        assert expect.get("cv").toString().equals(e.get("char").toString());
        assert expect.get("sv").toString().equals(e.get("short").toString());
        assert expect.get("nv").toString().equals(e.get("int").toString());
        assert expect.get("lv").toString().equals(e.get("long").toString());
        assert Float.compare((Float) expect.get("fv"), Float.parseFloat(e.get("float").toString())) == 0;
        assert Double.compare((Double) expect.get("dv"), Double.parseDouble(e.get("double").toString())) == 0;
        assert expect.get("s").equals(e.get("string"));
        assert ((BigDecimal) expect.get("mv")).setScale(4, BigDecimal.ROUND_HALF_DOWN).equals(new BigDecimal(e.get("decimal").toString()).setScale(4, BigDecimal.ROUND_HALF_DOWN));
        Date av1 = (Date) expect.get("av"), av2 = (Date) e.get("date");
        assert av1.getTime() / 1000 == av2.getTime() / 1000;
        Date iv1 = (Date) expect.get("iv"), iv2 = (Date) e.get("timestamp");
        assert iv1.getTime() / 1000 == iv2.getTime() / 1000;
        Time tv1 = (Time) expect.get("tv"), tv2 = (Time) e.get("time");
        assert String.valueOf(tv1).equals(String.valueOf(tv2));
        LocalDate ldv1 = (LocalDate) expect.get("ldv");
        Timestamp ldv2 = (Timestamp) e.get("LocalDate");
        assert ldv1.equals(ldv2.toLocalDateTime().toLocalDate());
        LocalDateTime ldtv1 = (LocalDateTime) expect.get("ldtv");
        Timestamp ldtv2 = (Timestamp) e.get("LocalDateTime");
        assert Timestamp.valueOf(ldtv1).getTime() / 1000 == ldtv2.getTime() / 1000;
        LocalTime ltv1 = (LocalTime) expect.get("ltv");
        Time ltv2 = (Time) e.get("LocalTime");
        assert String.valueOf(Time.valueOf(ltv1)).equals(String.valueOf(ltv2));
    }
}

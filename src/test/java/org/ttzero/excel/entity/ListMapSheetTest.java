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
import org.ttzero.excel.entity.e7.XMLWorksheetWriter;
import org.ttzero.excel.entity.style.Fill;
import org.ttzero.excel.entity.style.Horizontals;
import org.ttzero.excel.entity.style.NumFmt;
import org.ttzero.excel.entity.style.PatternType;
import org.ttzero.excel.entity.style.Styles;
import org.ttzero.excel.reader.Cell;
import org.ttzero.excel.reader.Drawings;
import org.ttzero.excel.reader.ExcelReader;
import org.ttzero.excel.reader.HeaderRow;
import org.ttzero.excel.reader.Row;
import org.ttzero.excel.util.ExtBufferedWriter;
import org.ttzero.excel.util.StringUtil;

import java.awt.Color;
import java.io.BufferedWriter;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.OutputStreamWriter;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.nio.ByteBuffer;
import java.nio.ByteOrder;
import java.nio.channels.SeekableByteChannel;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardOpenOption;
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
import java.util.HashSet;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.UUID;
import java.util.stream.Collectors;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;

/**
 * @author guanquan.wang at 2019-04-28 19:16
 */
public class ListMapSheetTest extends WorkbookTest {

    @Test public void testWrite() throws IOException {
        String fileName = "test map.xlsx";
        List<Map<String, Object>> expectList = createTestData();
        new Workbook()
            .addSheet(new ListMapSheet<>(expectList))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            List<Map<String, Object>> list = reader.sheet(0).dataRows().map(Row::toMap).collect(Collectors.toList());
            assertEquals(expectList.size(), list.size());
            for (int i = 0, len = expectList.size(); i < len; i++) {
                Map<String, Object> expect = expectList.get(i), e = list.get(i);
                assertEquals(expect, e);
            }
        }
    }

    @Test public void testAllType() throws IOException {
        String fileName = "test all type map.xlsx";
        List<Map<String, Object>> expectList = createAllTypeData();
        new Workbook()
            .setAutoSize(true)
            .addSheet(new ListMapSheet<>(expectList))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            List<Map<String, Object>> list = reader.sheet(0).dataRows().map(Row::toMap).collect(Collectors.toList());
            assertAllTypes(expectList, list);
        }
    }

    @Test public void testStyleDesign4Map() throws IOException {
        String fileName = "Map标识行样式.xlsx";
        List<Map<String, Object>> expectList = createAllTypeData(100);
        new Workbook()
                .addSheet(new ListMapSheet<>("Map", expectList).setStyleProcessor((map, style, sst) -> {
                    if ((Boolean) map.get("bv")) {
                        style = sst.modifyFill(style, new Fill(PatternType.solid, Color.green));
                    }
                    return style;
                }))
                .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            org.ttzero.excel.reader.Sheet sheet = reader.sheet(0).header(1);
            assertEquals("Map", sheet.getName());
            Iterator<org.ttzero.excel.reader.Row> iter = sheet.iterator();
            for (Map<String, Object> expect : expectList) {
                assertTrue(iter.hasNext());
                org.ttzero.excel.reader.Row row = iter.next();
                Map<String, Object> e = row.toMap();
                assertAllType(expect, e);

                boolean bv = (Boolean) expect.get("bv");
                Styles styles = row.getStyles();
                for (int i = row.getFirstColumnIndex(), j = row.getLastColumnIndex(); i < j; i++) {
                    int styleIndex = row.getCellStyle(i);
                    Fill fill = styles.getFill(styleIndex);
                    if (bv) {
                        assertTrue(fill != null && fill.getPatternType() == PatternType.solid && fill.getFgColor().equals(Color.green));
                    } else {
                        assertTrue(fill == null || fill.getPatternType() == PatternType.none);
                    }
                }
            }
        }
    }

    @Test public void testStyleDesign4Map2() throws IOException {
        String fileName = "Map标识行样式2.xlsx";
        List<Map<String, Object>> expectList = createAllTypeData(100);
        new Workbook()
            .addSheet(new ListMapSheet<>("Map", expectList
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
            assertEquals("Map", sheet.getName());
            org.ttzero.excel.reader.HeaderRow header = (HeaderRow) sheet.getHeader();
            assertEquals("boolean", header.get(0));
            assertEquals("char", header.get(1));
            assertEquals("short", header.get(2));
            assertEquals("int", header.get(3));
            assertEquals("long", header.get(4));
            assertEquals("LocalDateTime", header.get(5));
            assertEquals("LocalTime", header.get(6));
            Iterator<org.ttzero.excel.reader.Row> iter = sheet.iterator();
            for (Map<String, Object> expect : expectList) {
                assertTrue(iter.hasNext());
                org.ttzero.excel.reader.Row row = iter.next();
                Map<String, Object> e = row.toMap();

                assertEquals(expect.get("bv"), e.get("boolean"));
                assertEquals(expect.get("cv").toString(), e.get("char").toString());
                assertEquals(expect.get("sv").toString(), e.get("short").toString());
                assertEquals(expect.get("nv").toString(), e.get("int").toString());
                assertEquals(expect.get("lv").toString(), e.get("long").toString());
                LocalDateTime ldtv1 = (LocalDateTime) expect.get("ldtv");
                Timestamp ldtv2 = (Timestamp) e.get("LocalDateTime");
                assertEquals(Timestamp.valueOf(ldtv1).getTime() / 1000, ldtv2.getTime() / 1000);
                LocalTime ltv1 = (LocalTime) expect.get("ltv");
                Time ltv2 = (Time) e.get("LocalTime");
                assertEquals(String.valueOf(Time.valueOf(ltv1)), String.valueOf(ltv2));

                boolean bv = (Boolean) expect.get("bv");
                Styles styles = row.getStyles();
                for (int i = row.getFirstColumnIndex(), j = row.getLastColumnIndex(); i < j; i++) {
                    int styleIndex = row.getCellStyle(i);
                    Fill fill = styles.getFill(styleIndex);
                    if (bv) {
                        assertTrue(fill != null && fill.getPatternType() == PatternType.solid && fill.getFgColor().equals(Color.green));
                    } else {
                        assertTrue(fill == null || fill.getPatternType() == PatternType.none);
                    }
                }

                int styleIndex3 = row.getCellStyle(3), horizontals = styles.getHorizontal(styleIndex3);
                NumFmt numFmt = styles.getNumFmt(styleIndex3);
                assertEquals("¥0.00_);[Red]\\-¥0.00_);¥0_)", numFmt.getCode());
                if ((Integer) expect.get("nv") < 0) {
                    assertEquals(Horizontals.LEFT, horizontals);
                } else {
                    assertEquals(Horizontals.RIGHT, horizontals);
                }
            }
        }
    }

    @Test public void testHeaderColumn() throws IOException {
        String fileName = "test header column map.xlsx";
        List<Map<String, Object>> expectList = createAllTypeData();
        new Workbook()
            .addSheet(new ListMapSheet<>(expectList
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
            List<Map<String, Object>> list = reader.sheet(0).dataRows().map(Row::toMap).collect(Collectors.toList());
            assertEquals(expectList.size(), list.size());
            for (int i = 0, len = expectList.size(); i < len; i++) {
                Map<String, Object> expect = expectList.get(i), e = list.get(i);
                assertAllTypeFullKey(expect, e);
            }
        }
    }

    @Test public void testHeaderColumnBox() throws IOException {
        String fileName = "test header column box type map.xlsx";
        List<Map<String, Object>> expectList = createAllTypeData();
        new Workbook()
            .addSheet(new ListMapSheet<>(expectList
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
            assertEquals("Character", header.get(0));
            assertEquals("Short", header.get(1));
            assertEquals("Integer", header.get(2));
            assertEquals("Long", header.get(3));
            assertEquals("Float", header.get(4));
            assertEquals("Double", header.get(5));
            Iterator<org.ttzero.excel.reader.Row> iter = sheet.iterator();
            for (Map<String, Object> expect : expectList) {
                assertTrue(iter.hasNext());
                org.ttzero.excel.reader.Row row = iter.next();
                Map<String, Object> e = row.toMap();

                assertEquals(expect.get("cv").toString(), e.get("Character").toString());
                assertEquals(expect.get("sv").toString(), e.get("Short").toString());
                assertEquals(expect.get("nv").toString(), e.get("Integer").toString());
                assertEquals(expect.get("lv").toString(), e.get("Long").toString());
                assertEquals(Float.compare((Float) expect.get("fv"), Float.parseFloat(e.get("Float").toString())), 0);
                assertEquals(Double.compare((Double) expect.get("dv"), Double.parseDouble(e.get("Double").toString())), 0);
            }
        }
    }

    @Test public void testConstructor1() throws IOException {
        String fileName = "test list map sheet Constructor1.xlsx";
        new Workbook()
            .setAutoSize(true)
            .addSheet(new ListMapSheet<>())
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            assertEquals(reader.sheet(0).rows().count(), 0L);
        }
    }

    @Test public void testConstructor2() throws IOException {
        String fileName = "test list map sheet Constructor2.xlsx";
        List<Map<String, Object>> expectList = createTestData(10);
        new Workbook()
            .setAutoSize(true)
            .addSheet(new ListMapSheet<>("Map").setData(expectList))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            List<Map<String, Object>> list = reader.sheet(0).dataRows().map(Row::toMap).collect(Collectors.toList());
            assertEquals(expectList.size(), list.size());
            for (int i = 0, len = expectList.size(); i < len; i++) {
                Map<String, Object> expect = expectList.get(i), e = list.get(i);
                assertEquals(expect, e);
            }
        }
    }

    @Test public void testConstructor3() throws IOException {
        String fileName = "test list map sheet Constructor3.xlsx";
        List<Map<String, Object>> expectList = createAllTypeData(10);
        new Workbook()
            .setAutoSize(true)
            .addSheet(new ListMapSheet<>("Map"
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
            List<Map<String, Object>> list = reader.sheet(0).dataRows().map(Row::toMap).collect(Collectors.toList());
            assertEquals(expectList.size(), list.size());
            for (int i = 0, len = expectList.size(); i < len; i++) {
                Map<String, Object> expect = expectList.get(i), e = list.get(i);
                assertAllTypeFullKey(expect, e);
            }
        }
    }

    @Test public void testConstructor4() throws IOException {
        String fileName = "test list map sheet Constructor4.xlsx";
        List<Map<String, Object>> expectList = createAllTypeData(10);
        new Workbook()
            .setAutoSize(true)
            .addSheet(new ListMapSheet<>("Map", WaterMark.of(author)
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
            assertEquals(pictures.size(), 1);
            assertTrue(pictures.get(0).isBackground());

            List<Map<String, Object>> list = sheet.dataRows().map(Row::toMap).collect(Collectors.toList());
            assertEquals(expectList.size(), list.size());
            for (int i = 0, len = expectList.size(); i < len; i++) {
                Map<String, Object> expect = expectList.get(i), e = list.get(i);
                assertAllTypeFullKey(expect, e);
            }
        }
    }

    @Test public void testConstructor5() throws IOException {
        String fileName = "test list map sheet Constructor5.xlsx";
        List<Map<String, Object>> expectList = createAllTypeData(10);
        new Workbook()
            .setAutoSize(true)
            .addSheet(new ListMapSheet<>(expectList))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            org.ttzero.excel.reader.Sheet sheet = reader.sheet(0);
            List<Map<String, Object>> list = sheet.dataRows().map(Row::toMap).collect(Collectors.toList());
            assertEquals(expectList.size(), list.size());
            for (int i = 0, len = expectList.size(); i < len; i++) {
                Map<String, Object> expect = expectList.get(i), e = list.get(i);
                assertAllType(expect, e);
            }
        }
    }

    @Test public void testConstructor6() throws IOException {
        String fileName = "test list map sheet Constructor6.xlsx";
        List<Map<String, Object>> expectList = createAllTypeData(10);
        new Workbook()
            .setAutoSize(true)
            .addSheet(new ListMapSheet<>("Map", expectList))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            org.ttzero.excel.reader.Sheet sheet = reader.sheet(0);
            assertEquals("Map", sheet.getName());
            List<Map<String, Object>> list = sheet.dataRows().map(Row::toMap).collect(Collectors.toList());
            assertEquals(expectList.size(), list.size());
            for (int i = 0, len = expectList.size(); i < len; i++) {
                Map<String, Object> expect = expectList.get(i), e = list.get(i);
                assertAllType(expect, e);
            }
        }
    }

    @Test public void testConstructor8() throws IOException {
        String fileName = "test list map sheet Constructor8.xlsx";
        List<Map<String, Object>> expectList = createTestData(10);
        new Workbook()
            .setAutoSize(true)
            .addSheet(new ListMapSheet<>("MAP", expectList
                , new Column("ID", "id", int.class)
                , new Column("NAME", "name", String.class)))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            org.ttzero.excel.reader.Sheet sheet = reader.sheet(0);
            assertEquals("MAP", sheet.getName());
            List<Map<String, Object>> list = sheet.dataRows().map(Row::toMap).collect(Collectors.toList());
            assertEquals(expectList.size(), list.size());
            for (int i = 0, len = expectList.size(); i < len; i++) {
                Map<String, Object> expect = expectList.get(i), e = list.get(i);
                assertEquals(expect.get("id"), e.get("ID"));
                assertEquals(expect.get("name"), e.get("NAME"));
            }
        }
    }

    @Test public void testConstructor9() throws IOException {
        String fileName = "test list map sheet Constructor9.xlsx";
        List<Map<String, Object>> expectList = createTestData(10);
        new Workbook()
            .setAutoSize(true)
            .addSheet(new ListMapSheet<>(expectList
                , WaterMark.of(author)
                , new Column("ID", "id")
                , new Column("NAME", "name")))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            org.ttzero.excel.reader.Sheet sheet = reader.sheet(0);

            // Water Mark
            List<Drawings.Picture> pictures = sheet.listPictures();
            assertEquals(pictures.size(), 1);
            assertTrue(pictures.get(0).isBackground());

            List<Map<String, Object>> list = sheet.dataRows().map(Row::toMap).collect(Collectors.toList());
            assertEquals(expectList.size(), list.size());
            for (int i = 0, len = expectList.size(); i < len; i++) {
                Map<String, Object> expect = expectList.get(i), e = list.get(i);
                assertEquals(expect.get("id"), e.get("ID"));
                assertEquals(expect.get("name"), e.get("NAME"));
            }
        }
    }

    @Test public void testConstructor10() throws IOException {
        String fileName = "test list map sheet Constructor10.xlsx";
        List<Map<String, Object>> expectList = createTestData(10);
        new Workbook()
            .setAutoSize(true)
            .addSheet(new ListMapSheet<>("MAP"
                , expectList
                , WaterMark.of(author)
                , new Column("ID", "id", int.class)
                , new Column("NAME", "name", String.class)))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            org.ttzero.excel.reader.Sheet sheet = reader.sheet(0);
            assertEquals("MAP", sheet.getName());

            // Water Mark
            List<Drawings.Picture> pictures = sheet.listPictures();
            assertEquals(pictures.size(), 1);
            assertTrue(pictures.get(0).isBackground());

            List<Map<String, Object>> list = sheet.dataRows().map(Row::toMap).collect(Collectors.toList());
            assertEquals(expectList.size(), list.size());
            for (int i = 0, len = expectList.size(); i < len; i++) {
                Map<String, Object> expect = expectList.get(i), e = list.get(i);
                assertEquals(expect.get("id"), e.get("ID"));
                assertEquals(expect.get("name"), e.get("NAME"));
            }
        }
    }

    @Test public void testArray() throws IOException {
        String fileName = "ListMapSheet Array Map.xlsx";
        List<Map<String, Object>> expectList;
        Map<String, Object> data1 = new HashMap<>();
        data1.put("id", 1);
        data1.put("name", "abc");

        Map<String, Object> data2 = new HashMap<>();
        data2.put("id", 2);
        data2.put("name", "xyz");
        new Workbook()
            .addSheet(new ListMapSheet<>().setData(expectList = Arrays.asList(data1, data2)))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            org.ttzero.excel.reader.Sheet sheet = reader.sheet(0);
            List<Map<String, Object>> list = sheet.dataRows().map(Row::toMap).collect(Collectors.toList());
            assertEquals(expectList.size(), list.size());
            for (int i = 0, len = expectList.size(); i < len; i++) {
                Map<String, Object> expect = expectList.get(i), e = list.get(i);
                assertEquals(expect, e);
            }
        }
    }

    @Test public void testSingleList() throws IOException {
        String fileName = "ListMapSheet Single List Map.xlsx";
        List<Map<String, Object>> expectList;
        Map<String, Object> data = new HashMap<>();
        data.put("id", 1);
        data.put("name", "abc");

        new Workbook()
            .addSheet(new ListMapSheet<>().setData(expectList = Collections.singletonList(data)))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            org.ttzero.excel.reader.Sheet sheet = reader.sheet(0);
            List<Map<String, Object>> list = sheet.dataRows().map(Row::toMap).collect(Collectors.toList());
            assertEquals(expectList.size(), list.size());
            for (int i = 0, len = expectList.size(); i < len; i++) {
                Map<String, Object> expect = expectList.get(i), e = list.get(i);
                assertEquals(expect, e);
            }
        }
    }

    @Test public void testNullValue() throws IOException {
        String fileName = "test map null value.xlsx";
        List<Map<String, Object>> expectList = createNullTestData(10);
        new Workbook()
            .addSheet(new ListMapSheet<>(expectList))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            org.ttzero.excel.reader.Sheet sheet = reader.sheet(0);
            List<Map<String, Object>> list = sheet.dataRows().map(Row::toMap).collect(Collectors.toList());
            assertEquals(expectList.size(), list.size());
            for (int i = 0, len = expectList.size(); i < len; i++) {
                Map<String, Object> expect = expectList.get(i), e = list.get(i);
                assertEquals(expect.get("id"), e.get("id"));
                assertTrue(e.get("name") == null || StringUtil.isEmpty(e.get("name").toString()));
            }
        }
    }

    // Issue #93
    @Test public void testListMapSheet_93() throws IOException {
        String fileName = "Issue#93 List Map.xlsx";
        List<Map<String, Object>> expectList = new ArrayList<>();
        new Workbook().addSheet(new ListMapSheet<Object>() {
            private int i;
            @Override
            protected List<Map<String, Object>> more() {
                List<Map<String, Object>> list = i++ < 10 ? createAllTypeData(30) : null;
                if (list != null) expectList.addAll(list);
                return list;
            }
        }).writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            List<Map<String, Object>> list = reader.sheet(0).dataRows().map(Row::toMap).collect(Collectors.toList());
            assertAllTypes(expectList, list);
        }
    }

    @Test public void test_161() throws IOException {
        String fileName = "Issue#161.xlsx";
        List<Map<String, Object>> expectList = new ArrayList<>();
        new Workbook().addSheet(new ListMapSheet<Object>() {
            private int i = 0;
            @Override
            protected List<Map<String, Object>> more() {
                // Only write one row
                if (i++ > 0) return null;
                List<Map<String, Object>> list = new ArrayList<>();
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
            List<Map<String, Object>> list = reader.sheet(0).dataRows().map(Row::toMap).collect(Collectors.toList());
            assertEquals(list.size(), expectList.size());
            Map<String, Object> expect = expectList.get(0), e = list.get(0);
            assertEquals(expect.get("uuid"), e.get("uuid"));
            assertEquals(expect.get("hobbies").toString(), e.get("hobbies"));
            assertEquals(expect.get("sex"), e.get("sex"));
            assertEquals(expect.get("name"), e.get("name"));
            assertEquals(expect.get("age"), e.get("age"));
            LocalDateTime ldtv1 = (LocalDateTime) expect.get("createDate");
            Timestamp ldtv2 = (Timestamp) e.get("createDate");
            assertEquals(Timestamp.valueOf(ldtv1).getTime() / 1000, ldtv2.getTime() / 1000);
        }
    }

    @Test public void testWrapText() throws IOException {
        String fileName = "MAP WRAP TEXT.xlsx";
        List<Map<String, Object>> expectList = createTestData(10);
        new Workbook()
                .addSheet(new ListMapSheet<>(expectList
                    , new Column("ID", "id", int.class)
                    , new Column("NAME", "name", String.class).setWrapText(true)
                ))
                .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            List<Map<String, Object>> list = reader.sheet(0).dataRows().map(Row::toMap).collect(Collectors.toList());
            assertEquals(expectList.size(), list.size());
            for (int i = 0, len = expectList.size(); i < len; i++) {
                Map<String, Object> expect = expectList.get(i), e = list.get(i);
                assertEquals(expect.get("id"), e.get("ID"));
                assertEquals(expect.get("name"), e.get("NAME"));
            }
        }
    }

    @Test(expected = TooManyColumnsException.class) public void testOverLargeOrderColumn() throws IOException {
        new Workbook("test list map sheet Constructor8", author)
                .setAutoSize(true)
                .addSheet(new ListMapSheet<>("MAP", createTestData(10)
                        , new Column("ID", "id", int.class).setColIndex(9999999)
                        , new Column("NAME", "name", String.class)))
                .writeTo(defaultTestPath);
    }

    @Test public void test257() throws IOException {
        String fileName = "Issue#257.xlsx";
        List<Map<String, Object>> expectList = new ArrayList<>();
        expectList.add(new HashMap<String, Object>(){{put("sub1", "moban1");}});
        expectList.add(new HashMap<String, Object>(){{put("sub2", "moban2");}});
        expectList.add(new HashMap<String, Object>(){{put("sub3", "moban3");}});

        new Workbook().addSheet(new ListMapSheet<>(expectList
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
            assertEquals("ID", header.get(0));
            assertEquals("子表单", header.get(1));
            assertEquals("模板2", header.get(2));
            assertEquals("模板3", header.get(3));
            assertEquals("abc", header.get(4));
            assertEquals("模板2", header.get(5));
            assertEquals("xx", header.get(6));
            assertEquals("xyz", header.get(7));
            List<Map<String, Object>> list = reader.sheet(0).dataRows().map(Row::toMap).collect(Collectors.toList());
            assertEquals(expectList.size(), list.size());
            Map<String, Object> expect = expectList.get(0), e = list.get(0);
            assertEquals(expect.get("sub1"), e.get("子表单"));
            expect = expectList.get(1); e = list.get(1);
            assertEquals(expect.get("sub2"), e.get("模板2"));
            expect = expectList.get(2); e = list.get(2);
            assertEquals(expect.get("sub3"), e.get("模板3"));
            assertTrue(e.get("ID") == null || StringUtil.isEmpty(e.get("ID").toString()));
            assertTrue(e.get("abc") == null || StringUtil.isEmpty(e.get("abc").toString()));
            assertTrue(e.get("xx") == null || StringUtil.isEmpty(e.get("xx").toString()));
            assertTrue(e.get("xyz") == null || StringUtil.isEmpty(e.get("xyz").toString()));
        }
    }

    @Test public void testNullInListMap() throws IOException {
        String fileName = "Null in list map.xlsx";
        List<Map<String, Object>> expectList = createTestData(10);
        expectList.add(0, null);
        expectList.add(3, null);
        expectList.add(null);
        new Workbook().addSheet(new ListMapSheet<>(expectList)).writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            Iterator<org.ttzero.excel.reader.Row> iter = reader.sheet(0).header(1).iterator();
            for (Map<String, Object> expect : expectList) {
                assertTrue(iter.hasNext());
                Row row = iter.next();
                if (expect == null || expect.isEmpty()) {
                    assertTrue(row.isBlank());
                } else {
                    assertEquals(expect, row.toMap());
                }
            }
        }
    }

    @Test public void testLargeColumns() throws IOException {
        int len = 1436;
        List<Map<String, Object>> expectList = new ArrayList<>(len);
        for (int i = 0; i < len; i++) {
            Map<String, Object> map = new LinkedHashMap<>();
            for (int j = 0; j < 500; j++) {
                map.put("key" + j, getRandomString());
            }
            expectList.add(map);
        }

        new Workbook().addSheet(new ListMapSheet<>(expectList)).writeTo(defaultTestPath.resolve("large map.xlsx"));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("large map.xlsx"))) {
            List<Map<String, Object>> list = reader.sheet(0).dataRows().map(Row::toMap).collect(Collectors.toList());
            assertEquals(expectList.size(), list.size());
            for (int i = 0; i < len; i++) {
                Map<String, Object> expect = expectList.get(i), e = list.get(i);
                assertEquals(expect, e);
            }
        }
    }

    @Test public void testSpecifyRowWrite() throws IOException {
        List<Map<String, Object>> list = createTestData(10);
        new Workbook().setAutoSize(true)
            .addSheet(new ListMapSheet<>(list).setStartRowIndex(5))
            .writeTo(defaultTestPath.resolve("test specify row 5 ListMapSheet.xlsx"));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("test specify row 5 ListMapSheet.xlsx"))) {
            List<Map<String, Object>> readList = reader.sheet(0).header(5).rows().map(Row::toMap).collect(Collectors.toList());
            assertEquals(list.size(), readList.size());
            for (int i = 0, len = list.size(); i < len; i++) {
                Map<String, Object> r = readList.get(i), w = list.get(i);
                assertEquals(r, w);
            }
        }
    }

    @Test public void testSpecifyRowStayA1Write() throws IOException {
        List<Map<String, Object>> list = createTestData(10);
        new Workbook().setAutoSize(true)
            .addSheet(new ListMapSheet<>(list).setStartRowIndex(5, false))
            .writeTo(defaultTestPath.resolve("test specify row 5 stay A1 ListMapSheet.xlsx"));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("test specify row 5 stay A1 ListMapSheet.xlsx"))) {
            List<Map<String, Object>> readList = reader.sheet(0).header(5).rows().map(Row::toMap).collect(Collectors.toList());
            assertEquals(list.size(), readList.size());
            for (int i = 0, len = list.size(); i < len; i++) {
                Map<String, Object> r = readList.get(i), w = list.get(i);
                assertEquals(r, w);
            }
        }
    }

    @Test public void testSpecifyRowAndColWrite() throws IOException {
        List<Map<String, Object>> list = createTestData(10);
        new Workbook().setAutoSize(true)
            .addSheet(new ListMapSheet<>(list
                , new Column("id").setColIndex(3)
                , new Column("name").setColIndex(4))
                .setStartRowIndex(5))
            .writeTo(defaultTestPath.resolve("test specify row and col ListMapSheet.xlsx"));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("test specify row and col ListMapSheet.xlsx"))) {
            List<Map<String, Object>> readList = reader.sheet(0).header(5).rows().map(Row::toMap).collect(Collectors.toList());
            assertEquals(list.size(), readList.size());
            for (int i = 0, len = list.size(); i < len; i++) {
                Map<String, Object> r = readList.get(i), w = list.get(i);
                assertEquals(r.size(), w.size());
                assertEquals(r.get("id"), w.get("id"));
                assertEquals(r.get("name"), w.get("name"));
            }
        }
    }

    @Test public void testSpecifyRowAndColStayA1Write() throws IOException {
        List<Map<String, Object>> list = createTestData(10);
        new Workbook().setAutoSize(true)
            .addSheet(new ListMapSheet<>(list
                , new Column("id").setColIndex(3)
                , new Column("name").setColIndex(4))
                .setStartRowIndex(5, false))
            .writeTo(defaultTestPath.resolve("test specify row and cel stay A1 ListMapSheet.xlsx"));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("test specify row and cel stay A1 ListMapSheet.xlsx"))) {
            List<Map<String, Object>> readList = reader.sheet(0).header(5).rows().map(Row::toMap).collect(Collectors.toList());
            assertEquals(list.size(), readList.size());
            for (int i = 0, len = list.size(); i < len; i++) {
                Map<String, Object> r = readList.get(i), w = list.get(i);
                assertEquals(r, w);
            }
        }
    }

    @Test public void testAppendKeyMap() throws IOException {
        final String fileName = "append key map.xlsx";
        new Workbook().addSheet(new AppendKeyMapSheet() { // <- 使用自定义数据源
            int page = 1;

            @Override
            protected List<Map<String, Object>> more() {
                return getRows(page++);
            }
        }.setSheetWriter(new AppendHeaderWorksheetWriter())) // <- 使用自定义输出协议
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            Iterator<org.ttzero.excel.reader.Row> iter = reader.sheet(0).iterator();
            assertTrue(iter.hasNext());
            org.ttzero.excel.reader.Row row = iter.next();

        }
    }

    @Test public void testDataSupplier() throws IOException {
        final String fileName = "list map data supplier.xlsx";
        List<Map<String, Object>> expectList = new ArrayList<>(100);
        new Workbook()
            .addSheet(new ListMapSheet<>().setData((i, lastOne) -> {
                if (i >= 100) return null;
                List<Map<String, Object>> sub = createTestData(10);
                expectList.addAll(sub);
                return sub;
            }))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            List<Map<String, Object>> list =  reader.sheet(0).dataRows().map(Row::toMap).collect(Collectors.toList());
            assertEquals(expectList.size(), list.size());
            for (int i = 0, len = expectList.size(); i < len; i++) {
                Map<String, Object> expect = expectList.get(i), e = list.get(i);
                assertEquals(expect, e);
            }
        }
    }

    static List<Map<String, Object>> getRows(int page) {
        if (page > 127 - 67) return null;
        List<Map<String, Object>> rows = new ArrayList<>();//模拟hbase中的数据，这里查询了hbase
        Map<String, Object> map = new HashMap<>();
        map.put("A", "a" + page);
        map.put("B", "b" + page);
        String key = String.valueOf((char) ('B' + page));
        map.put(key, key + page);
        rows.add(map);
        return rows;
    }

    public static class AppendKeyMapSheet extends ListMapSheet<Object> {
        // 保存已存在中列
        Set<String> existsKeys = new HashSet<>();

        @Override
        public Column[] getHeaderColumns() {
            Column[] columns = super.getHeaderColumns();
            for (Column col : columns) existsKeys.add(col.name);
            return columns;
        }

        @Override
        protected void resetBlockData() {
            if (!eof && left() < rowBlock.capacity()) append();
            int end = getEndIndex(), len;
            for (; start < end; rows++, start++) {
                org.ttzero.excel.entity.Row row = rowBlock.next();
                row.index = rows;
                row.height = getRowHeight();
                Map<String, Object> rowDate = data.get(start);
                boolean isNull = rowDate == null;

                if (!isNull) {
                    // 检查是否有不存在的key
                    for (String k : rowDate.keySet()) {
                        if (!existsKeys.contains(k)) {
                            Column col = new Column(k, k), pre = columns[columns.length - 1].getTail(); // <- 需要判断NPE
                            col.colIndex = pre.colIndex + 1;
                            col.realColIndex = pre.realColIndex + 1;
                            col.styles = getWorkbook().getStyles();
                            // 扩容并追加到末尾
                            columns = Arrays.copyOf(columns, columns.length + 1);
                            columns[columns.length - 1] = col;
                            existsKeys.add(k);
                        }
                    }
                }
                len = columns.length;

                Cell[] cells = row.realloc(len);
                for (int i = 0; i < len; i++) {
                    Column hc = columns[i];
                    Object e = !isNull ? rowDate.get(hc.key) : null;
                    // Clear cells
                    Cell cell = cells[i];
                    cell.clear();

                    cellValueAndStyle.reset(row, cell, e, hc);
                }
            }
        }
    }

    public static class AppendHeaderWorksheetWriter extends XMLWorksheetWriter {
        // 记录body的位置
        long position = 0L;

        @Override
        protected void beforeSheetData(boolean nonHeader) throws IOException {
            super.beforeSheetData(nonHeader);
            bw.flush(); // 刷新流

            // 获取当前position，从position以后开始写实际的body
            position = Files.size(workSheetPath.resolve(sheet.getFileName()));
        }

        @Override
        public void close() {
            super.close();

            XMLWorksheetWriter _writer = new XMLWorksheetWriter(sheet) {
                @Override
                protected boolean hasMedia() {
                    return false;
                }
            };
            Class<XMLWorksheetWriter> clazz = XMLWorksheetWriter.class;
            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            try {
                Field totalRowsField = clazz.getDeclaredField("totalRows");
                totalRowsField.setAccessible(true);
                totalRowsField.set(_writer, totalRows);
                Field startRowField = clazz.getDeclaredField("startRow");
                startRowField.setAccessible(true);
                startRowField.set(_writer, startRow - columns[0].subColumnSize());
                Field startHeaderRowField = clazz.getDeclaredField("startHeaderRow");
                startHeaderRowField.setAccessible(true);
                startHeaderRowField.set(_writer, startHeaderRow);
                Field includeAutoWidthField = clazz.getDeclaredField("includeAutoWidth");
                includeAutoWidthField.setAccessible(true);
                includeAutoWidthField.set(_writer, includeAutoWidth);
                Field stylesField = clazz.getDeclaredField("styles");
                stylesField.setAccessible(true);
                stylesField.set(_writer, styles);
                Field bwField = clazz.getDeclaredField("bw");
                bwField.setAccessible(true);
                BufferedWriter bw = new ExtBufferedWriter(new OutputStreamWriter(baos, StandardCharsets.UTF_8));
                bwField.set(_writer, bw);
                // 重写col
                Method writeBeforeMethod = clazz.getDeclaredMethod("writeBefore");
                writeBeforeMethod.setAccessible(true);
                writeBeforeMethod.invoke(_writer);

                // 重写表头
                Method beforeSheetDataMethod = clazz.getDeclaredMethod("beforeSheetData", boolean.class);
                beforeSheetDataMethod.setAccessible(true);
                beforeSheetDataMethod.invoke(_writer, sheet.getNonHeader() == 1);

                bw.close();
            } catch (NoSuchFieldException | IllegalAccessException | InvocationTargetException | NoSuchMethodException | IOException e) {
                e.printStackTrace();
            }

            try {
                Path currentPath = workSheetPath.resolve(sheet.getFileName());
                String fileName = currentPath.getFileName().toString();
                // 创建临时文件
                Path tmpPath = Files.createFile(workSheetPath.resolve(fileName + "_cp"));
                // 将新表头复制到临时文件中
                try (SeekableByteChannel channel = Files.newByteChannel(tmpPath, StandardOpenOption.WRITE, StandardOpenOption.READ)) {
                    ByteBuffer buffer = ByteBuffer.wrap(baos.toByteArray());
                    buffer.order(ByteOrder.LITTLE_ENDIAN);
                    channel.write(buffer);
                }

                // 将Body复制到临时文件中
                try (InputStream is = Files.newInputStream(currentPath);
                     OutputStream os = Files.newOutputStream(tmpPath, StandardOpenOption.APPEND)) {
                    is.skip(position); // <- 跳到body处

                    byte[] bytes = new byte[4096];
                    int n;
                    while ((n = is.read(bytes)) > 0) {
                        os.write(bytes, 0, n);
                    }
                }

                // 替换现有文件
                Files.delete(currentPath);
                tmpPath.toFile().renameTo(currentPath.toFile());
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    public static List<Map<String, Object>> createTestData() {
        int size = random.nextInt(100) + 1;
        return createTestData(size);
    }

    public static List<Map<String, Object>> createTestData(int size) {
        List<Map<String, Object>> list = new ArrayList<>(size);
        for (int i = 0; i < size; i++) {
            Map<String, Object> map = new LinkedHashMap<>();
            map.put("id", random.nextInt());
            map.put("name", getRandomString());
            list.add(map);
        }
        return list;
    }

    public static List<Map<String, Object>> createAllTypeData() {
        int size = random.nextInt(100) + 1;
        return createAllTypeData(size);
    }

    public static List<Map<String, Object>> createAllTypeData(int size) {
        List<Map<String, Object>> list = new ArrayList<>(size);
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

    public static List<Map<String, Object>> createNullTestData(int size) {
        List<Map<String, Object>> list = new ArrayList<>(size);
        for (int i = 0; i < size; i++) {
            Map<String, Object> map = new LinkedHashMap<>();
            map.put("id", random.nextInt());
            map.put("name", null);
            list.add(map);
        }
        return list;
    }

    static void assertAllTypes(List<Map<String, Object>> expectList, List<Map<String, Object>> list) {
        assertEquals(expectList.size(), list.size());
        for (int i = 0, len = expectList.size(); i < len; i++) {
            Map<String, Object> expect = expectList.get(i), e = list.get(i);
            assertAllType(expect, e);
        }
    }

    static void assertAllType(Map<String, Object> expect, Map<String, Object> e) {
        assertEquals(expect.size(), e.size());
        assertEquals(expect.get("bv"), e.get("bv"));
        assertEquals(expect.get("cv").toString(), e.get("cv").toString());
        assertEquals(expect.get("sv").toString(), e.get("sv").toString());
        assertEquals(expect.get("nv").toString(), e.get("nv").toString());
        assertEquals(expect.get("lv").toString(), e.get("lv").toString());
        assertEquals(Float.compare((Float) expect.get("fv"), Float.parseFloat(e.get("fv").toString())), 0);
        assertEquals(Double.compare((Double) expect.get("dv"), Double.parseDouble(e.get("dv").toString())), 0);
        assertEquals(expect.get("s"), e.get("s"));
        assertEquals(((BigDecimal) expect.get("mv")).setScale(4, RoundingMode.HALF_DOWN), new BigDecimal(e.get("mv").toString()).setScale(4, RoundingMode.HALF_DOWN));
        Date av1 = (Date) expect.get("av"), av2 = (Date) e.get("av");
        assertEquals(av1.getTime() / 1000, av2.getTime() / 1000);
        Date iv1 = (Date) expect.get("iv"), iv2 = (Date) e.get("iv");
        assertEquals(iv1.getTime() / 1000, iv2.getTime() / 1000);
        Time tv1 = (Time) expect.get("tv"), tv2 = (Time) e.get("tv");
        assertEquals(String.valueOf(tv1), String.valueOf(tv2));
        LocalDate ldv1 = (LocalDate) expect.get("ldv");
        Timestamp ldv2 = (Timestamp) e.get("ldv");
        assertEquals(ldv1, ldv2.toLocalDateTime().toLocalDate());
        LocalDateTime ldtv1 = (LocalDateTime) expect.get("ldtv");
        Timestamp ldtv2 = (Timestamp) e.get("ldtv");
        assertEquals(Timestamp.valueOf(ldtv1).getTime() / 1000, ldtv2.getTime() / 1000);
        LocalTime ltv1 = (LocalTime) expect.get("ltv");
        Time ltv2 = (Time) e.get("ltv");
        assertEquals(String.valueOf(Time.valueOf(ltv1)), String.valueOf(ltv2));
    }

    static void assertAllTypeFullKey(Map<String, Object> expect, Map<String, Object> e) {
        assertEquals(expect.size(), e.size());
        assertEquals(expect.get("bv"), e.get("boolean"));
        assertEquals(expect.get("cv").toString(), e.get("char").toString());
        assertEquals(expect.get("sv").toString(), e.get("short").toString());
        assertEquals(expect.get("nv").toString(), e.get("int").toString());
        assertEquals(expect.get("lv").toString(), e.get("long").toString());
        assertEquals(Float.compare((Float) expect.get("fv"), Float.parseFloat(e.get("float").toString())), 0);
        assertEquals(Double.compare((Double) expect.get("dv"), Double.parseDouble(e.get("double").toString())), 0);
        assertEquals(expect.get("s"), e.get("string"));
        assertEquals(((BigDecimal) expect.get("mv")).setScale(4, RoundingMode.HALF_DOWN), new BigDecimal(e.get("decimal").toString()).setScale(4, RoundingMode.HALF_DOWN));
        Date av1 = (Date) expect.get("av"), av2 = (Date) e.get("date");
        assertEquals(av1.getTime() / 1000, av2.getTime() / 1000);
        Date iv1 = (Date) expect.get("iv"), iv2 = (Date) e.get("timestamp");
        assertEquals(iv1.getTime() / 1000, iv2.getTime() / 1000);
        Time tv1 = (Time) expect.get("tv"), tv2 = (Time) e.get("time");
        assertEquals(String.valueOf(tv1), String.valueOf(tv2));
        LocalDate ldv1 = (LocalDate) expect.get("ldv");
        Timestamp ldv2 = (Timestamp) e.get("LocalDate");
        assertEquals(ldv1, ldv2.toLocalDateTime().toLocalDate());
        LocalDateTime ldtv1 = (LocalDateTime) expect.get("ldtv");
        Timestamp ldtv2 = (Timestamp) e.get("LocalDateTime");
        assertEquals(Timestamp.valueOf(ldtv1).getTime() / 1000, ldtv2.getTime() / 1000);
        LocalTime ltv1 = (LocalTime) expect.get("ltv");
        Time ltv2 = (Time) e.get("LocalTime");
        assertEquals(String.valueOf(Time.valueOf(ltv1)), String.valueOf(ltv2));
    }
}

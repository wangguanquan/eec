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
import org.ttzero.excel.Print;
import org.ttzero.excel.entity.style.Fill;
import org.ttzero.excel.entity.style.Horizontals;
import org.ttzero.excel.entity.style.PatternType;
import org.ttzero.excel.entity.style.Styles;
import org.ttzero.excel.reader.ExcelReader;

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
import java.util.List;
import java.util.Map;
import java.util.UUID;

/**
 * @author guanquan.wang at 2019-04-28 19:16
 */
public class ListMapSheetTest extends WorkbookTest {

    @Test public void testWrite() throws IOException {
        new Workbook("test map", author)
            .watch(Print::println)
            .addSheet(createTestData())
            .writeTo(defaultTestPath);
    }

    @Test public void testAllType() throws IOException {
        new Workbook("test all type map", author)
            .watch(Print::println)
            .setAutoSize(true)
            .addSheet(createAllTypeData())
            .writeTo(defaultTestPath);
    }

    @Test public void testStyleDesign4Map() throws IOException {
        new Workbook("Map标识行样式", author)
                .addSheet(new ListMapSheet("Map", createAllTypeData(100)).setStyleProcessor((map, style, sst)->{
                    if ((Boolean) map.get("bv")) {
                        style = Styles.clearFill(style) | sst.addFill(new Fill(PatternType.solid, Color.green));
                    }
                    return style;
                }))
                .writeTo(defaultTestPath);
    }

    @Test public void testStyleDesign4Map2() throws IOException {
        new Workbook("Map标识行样式2", author)
            .addSheet(new ListMapSheet("Map", createAllTypeData(100)
                , new Column("boolean", "bv", boolean.class)
                , new Column("char", "cv", char.class)
                , new Column("short", "sv", short.class)
                , new Column("int", "nv", int.class).setStyleProcessor((n,s,sst) -> ((int) n) < 0 ? Styles.clearHorizontal(s) | Horizontals.LEFT : s).setNumFmt("¥0.00_);[Red]-¥0.00_);¥0_)")
                , new Column("long", "lv", long.class)
                , new Column("LocalDateTime", "ldtv", LocalDateTime.class)
                , new Column("LocalTime", "ltv", LocalTime.class)).setStyleProcessor((map, style, sst)->{
                if ((Boolean) map.get("bv")) {
                    style = Styles.clearFill(style) | sst.addFill(new Fill(PatternType.solid, Color.green));
                }
                return style;
            }))
            .writeTo(defaultTestPath);
    }

    @Test public void testHeaderColumn() throws IOException {
        new Workbook("test header column map", author)
            .watch(Print::println)
            .addSheet(createAllTypeData()
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
            )
            .writeTo(defaultTestPath);
    }

    @Test public void testHeaderColumnBox() throws IOException {
        new Workbook("test header column box type map", author)
            .watch(Print::println)
            .addSheet(createAllTypeData()
                , new Column("Character", "cv", Character.class)
                , new Column("Short", "sv", Short.class)
                , new Column("Integer", "nv", Integer.class)
                , new Column("Long", "lv", Long.class)
                , new Column("Float", "fv", Float.class)
                , new Column("Double", "dv", Double.class)
            )
            .writeTo(defaultTestPath);
    }

    @Test public void testConstructor1() throws IOException {
        new Workbook("test list map sheet Constructor1", author)
            .watch(Print::println)
            .setAutoSize(true)
            .addSheet(new ListMapSheet())
            .writeTo(defaultTestPath);
    }

    @Test public void testConstructor2() throws IOException {
        new Workbook("test list map sheet Constructor2", author)
            .watch(Print::println)
            .setAutoSize(true)
            .addSheet(new ListMapSheet("Map").setData(createTestData(10)))
            .writeTo(defaultTestPath);
    }

    @Test public void testConstructor3() throws IOException {
        new Workbook("test list map sheet Constructor3", author)
            .watch(Print::println)
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
            )
                .setData(createAllTypeData(10)))
            .writeTo(defaultTestPath);
    }

    @Test public void testConstructor4() throws IOException {
        new Workbook("test list map sheet Constructor4", author)
            .watch(Print::println)
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
                .setData(createAllTypeData(10)))
            .writeTo(defaultTestPath);
    }

    @Test public void testConstructor5() throws IOException {
        new Workbook("test list map sheet Constructor5", author)
            .watch(Print::println)
            .setAutoSize(true)
            .addSheet(new ListMapSheet(createAllTypeData(10)))
            .writeTo(defaultTestPath);
    }

    @Test public void testConstructor6() throws IOException {
        new Workbook("test list map sheet Constructor6", author)
            .watch(Print::println)
            .setAutoSize(true)
            .addSheet(new ListMapSheet("Map", createAllTypeData(10)))
            .writeTo(defaultTestPath);
    }

    @Test public void testConstructor7() throws IOException {
        new Workbook("test list map sheet Constructor7", author)
            .watch(Print::println)
            .setAutoSize(true)
            .addSheet(new ListMapSheet(createAllTypeData(10)
                , new Column("Character", "cv", Character.class)
                , new Column("Short", "sv", Short.class)
                , new Column("Integer", "nv", Integer.class)
                , new Column("Long", "lv", Long.class)
                , new Column("Float", "fv", Float.class)
                , new Column("Double", "dv", Double.class)))
            .writeTo(defaultTestPath);
    }

    @Test public void testConstructor8() throws IOException {
        new Workbook("test list map sheet Constructor8", author)
            .watch(Print::println)
            .setAutoSize(true)
            .addSheet(new ListMapSheet("MAP", createTestData(10)
                , new Column("ID", "id", int.class)
                , new Column("NAME", "name", String.class)))
            .writeTo(defaultTestPath);
    }

    @Test public void testConstructor9() throws IOException {
        new Workbook("test list map sheet Constructor9", author)
            .watch(Print::println)
            .setAutoSize(true)
            .addSheet(new ListMapSheet(createTestData(10)
                , WaterMark.of(author)
                , new Column("ID", "id")
                , new Column("NAME", "name")))
            .writeTo(defaultTestPath);
    }

    @Test public void testConstructor10() throws IOException {
        new Workbook("test list map sheet Constructor10", author)
            .watch(Print::println)
            .setAutoSize(true)
            .addSheet(new ListMapSheet("MAP"
                , createTestData(10)
                , WaterMark.of(author)
                , new Column("ID", "id", int.class)
                , new Column("NAME", "name", String.class)))
            .writeTo(defaultTestPath);
    }

    @Test public void testArray() throws IOException {
        Map<String, Object> data1 = new HashMap<>();
        data1.put("id", 1);
        data1.put("name", "abc");

        Map<String, Object> data2 = new HashMap<>();
        data2.put("id", 2);
        data2.put("name", "xyz");
        new Workbook()
            .watch(Print::println)
            .addSheet(new ListMapSheet().setData(Arrays.asList(data1, data2)))
            .writeTo(defaultTestPath);
    }

    @Test public void testSingleList() throws IOException {
        Map<String, Object> data = new HashMap<>();
        data.put("id", 1);
        data.put("name", "abc");

        new Workbook()
            .watch(Print::println)
            .addSheet(new ListMapSheet().setData(Collections.singletonList(data)))
            .writeTo(defaultTestPath);
    }

    @Test public void testNullValue() throws IOException {
        new Workbook("test map null value", author)
            .watch(Print::println)
            .addSheet(createNullTestData(10))
            .writeTo(defaultTestPath);
    }

    // Issue #93
    @Test public void testListMapSheet_93() throws IOException {
        new Workbook("Issue#93 List Map").addSheet(new ListMapSheet() {
            private int i;
            @Override
            protected List<Map<String, ?>> more() {
                return i++ < 10 ? createAllTypeData(30) : null;
            }
        }).writeTo(defaultTestPath);
    }

    @Test public void test_161() throws IOException {
        new Workbook(("Issue#161")).addSheet(new ListMapSheet() {
            private int i = 0;
            @Override
            protected List<Map<String, ?>> more() {
                // Only write one row
                if (i++ > 0) return null;
                List<Map<String, ?>> list = new ArrayList<>();
                Map<String, Object> map = new HashMap<>();
                map.put("a0172da4c398047aeac758ecd4a799b71", UUID.randomUUID().toString());
                map.put("hobbies", new ArrayList<String>() {{
                    add("张");
                    add("李");
                }});
                map.put("sex", "男");
                final int len = 4095;
                StringBuilder buf = new StringBuilder(len);
                for (int i = 0; i < len; i++) {
                    buf.append('a');
                }
                // java.nio.BufferOverflowException occur when the cell value length large than 2045
                map.put("name", buf.toString());
                map.put("age", 24);
                map.put("createDate", new Date(1535444725000L).toInstant().atOffset(ZoneOffset.of("+8")).toLocalDateTime());

                list.add(map);
                return list;
            }
        }).writeTo(defaultTestPath);
    }

    @Test public void testWrapText() throws IOException {
        new Workbook("MAP WRAP TEXT", author)
                .addSheet(createTestData(10)
                    , new Column("ID", "id", int.class)
                    , new Column("NAME", "name", String.class).setWrapText(true)
                )
                .writeTo(defaultTestPath);
    }

    @Test public void testOverLargeOrderColumn() throws IOException {
        try {
            new Workbook("test list map sheet Constructor8", author)
                    .watch(Print::println)
                    .setAutoSize(true)
                    .addSheet(new ListMapSheet("MAP", createTestData(10)
                            , new Column("ID", "id", int.class).setColIndex(9999999)
                            , new Column("NAME", "name", String.class)))
                    .writeTo(defaultTestPath);
            assert false;
        } catch (TooManyColumnsException e) {
            assert true;
        }
    }

    @Test public void test257() throws IOException {
        List<Map<String, ?>> list = new ArrayList<>();
        list.add(new HashMap<String, String>(){{put("sub1", "moban1");}});
        list.add(new HashMap<String, String>(){{put("sub2", "moban2");}});
        list.add(new HashMap<String, String>(){{put("sub3", "moban3");}});

        new Workbook("Issue#257").addSheet(new ListMapSheet(list
                , new Column("ID", "id")
                , new Column("子表单", "sub1")
                , new Column("模板2", "sub2")
                , new Column("模板3", "sub3")
                , new Column("abc", "abc")
                , new Column("模板2", "sub2")
                , new Column("xx", "xx")
                , new Column("xyz", "xyz")
        )).writeTo(defaultTestPath);
    }

    @Test public void testNullInListMap() throws IOException {
        List<Map<String, ?>> list = createTestData(10);
        list.add(0, null);
        list.add(3, null);
        list.add(null);
        new Workbook("Null in list map").addSheet(new ListSheet<>(list)).writeTo(defaultTestPath);
    }

    @Test public void testLargeColumns() throws IOException {
        int len = 1436;
        List<Map<String, ?>> list = new ArrayList<>(len);
        for (int i = 0; i < len; i++) {
            Map<String, String> map = new HashMap<>();
            for (int j = 0; j < 500; j++) {
                map.put("key" + j, getRandomString());
            }
            list.add(map);
        }

        new Workbook().addSheet(new ListMapSheet(list)).writeTo(defaultTestPath.resolve("large map.xlsx"));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("large map.xlsx"))) {
            assert "A1:SF1437".equals(reader.sheet(0).getDimension().toString());
            long count = reader.sheet(0).rows().filter(row -> !row.isEmpty()).count();
            assert count == 1437;
        }
    }


    public static List<Map<String, ?>> createTestData() {
        int size = random.nextInt(100) + 1;
        return createTestData(size);
    }

    public static List<Map<String, ?>> createTestData(int size) {
        List<Map<String, ?>> list = new ArrayList<>(size);
        for (int i = 0; i < size; i++) {
            Map<String, Object> map = new HashMap<>();
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
            Map<String, Object> map = new HashMap<>();
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
            Map<String, Object> map = new HashMap<>();
            map.put("id", random.nextInt());
            map.put("name", null);
            list.add(map);
        }
        return list;
    }
}

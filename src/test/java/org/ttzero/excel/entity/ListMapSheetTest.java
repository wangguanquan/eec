/*
 * Copyright (c) 2019, guanquan.wang@yandex.com All Rights Reserved.
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

import java.io.IOException;
import java.math.BigDecimal;
import java.sql.Time;
import java.sql.Timestamp;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Create by guanquan.wang at 2019-04-28 19:16
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

    @Test public void testHeaderColumn() throws IOException {
        new Workbook("test header column map", author)
            .watch(Print::println)
            .addSheet(createAllTypeData()
                , new Sheet.Column("boolean", "bv", boolean.class)
                , new Sheet.Column("char", "cv", char.class)
                , new Sheet.Column("short", "sv", short.class)
                , new Sheet.Column("int", "nv", int.class)
                , new Sheet.Column("long", "lv", long.class)
                , new Sheet.Column("float", "fv", float.class)
                , new Sheet.Column("double", "dv", double.class)
                , new Sheet.Column("string", "s", String.class)
                , new Sheet.Column("decimal", "mv", BigDecimal.class)
                , new Sheet.Column("date", "av", Date.class)
                , new Sheet.Column("timestamp", "iv", Timestamp.class)
                , new Sheet.Column("time", "tv", Time.class)
                , new Sheet.Column("LocalDate", "ldv", LocalDate.class)
                , new Sheet.Column("LocalDateTime", "ldtv", LocalDateTime.class)
                , new Sheet.Column("LocalTime", "ltv", LocalTime.class)
            )
            .writeTo(defaultTestPath);
    }

    @Test public void testHeaderColumnBox() throws IOException {
        new Workbook("test header column box type map", author)
            .watch(Print::println)
            .addSheet(createAllTypeData()
                , new Sheet.Column("Character", "cv", Character.class)
                , new Sheet.Column("Short", "sv", Short.class)
                , new Sheet.Column("Integer", "nv", Integer.class)
                , new Sheet.Column("Long", "lv", Long.class)
                , new Sheet.Column("Float", "fv", Float.class)
                , new Sheet.Column("Double", "dv", Double.class)
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
                , new Sheet.Column("boolean", "bv", boolean.class)
                , new Sheet.Column("char", "cv", char.class)
                , new Sheet.Column("short", "sv", short.class)
                , new Sheet.Column("int", "nv", int.class)
                , new Sheet.Column("long", "lv", long.class)
                , new Sheet.Column("float", "fv", float.class)
                , new Sheet.Column("double", "dv", double.class)
                , new Sheet.Column("string", "s", String.class)
                , new Sheet.Column("decimal", "mv", BigDecimal.class)
                , new Sheet.Column("date", "av", Date.class)
                , new Sheet.Column("timestamp", "iv", Timestamp.class)
                , new Sheet.Column("time", "tv", Time.class)
                , new Sheet.Column("LocalDate", "ldv", LocalDate.class)
                , new Sheet.Column("LocalDateTime", "ldtv", LocalDateTime.class)
                , new Sheet.Column("LocalTime", "ltv", LocalTime.class)
            )
                .setData(createAllTypeData(10)))
            .writeTo(defaultTestPath);
    }

    @Test public void testConstructor4() throws IOException {
        new Workbook("test list map sheet Constructor4", author)
            .watch(Print::println)
            .setAutoSize(true)
            .addSheet(new ListMapSheet("Map", WaterMark.of(author)
                , new Sheet.Column("boolean", "bv", boolean.class)
                , new Sheet.Column("char", "cv", char.class)
                , new Sheet.Column("short", "sv", short.class)
                , new Sheet.Column("int", "nv", int.class)
                , new Sheet.Column("long", "lv", long.class)
                , new Sheet.Column("float", "fv", float.class)
                , new Sheet.Column("double", "dv", double.class)
                , new Sheet.Column("string", "s", String.class)
                , new Sheet.Column("decimal", "mv", BigDecimal.class)
                , new Sheet.Column("date", "av", Date.class)
                , new Sheet.Column("timestamp", "iv", Timestamp.class)
                , new Sheet.Column("time", "tv", Time.class)
                , new Sheet.Column("LocalDate", "ldv", LocalDate.class)
                , new Sheet.Column("LocalDateTime", "ldtv", LocalDateTime.class)
                , new Sheet.Column("LocalTime", "ltv", LocalTime.class))
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
                , new Sheet.Column("Character", "cv", Character.class)
                , new Sheet.Column("Short", "sv", Short.class)
                , new Sheet.Column("Integer", "nv", Integer.class)
                , new Sheet.Column("Long", "lv", Long.class)
                , new Sheet.Column("Float", "fv", Float.class)
                , new Sheet.Column("Double", "dv", Double.class)))
            .writeTo(defaultTestPath);
    }

    @Test public void testConstructor8() throws IOException {
        new Workbook("test list map sheet Constructor8", author)
            .watch(Print::println)
            .setAutoSize(true)
            .addSheet(new ListMapSheet("MAP", createTestData(10)
                , new Sheet.Column("ID", "id", int.class)
                , new Sheet.Column("NAME", "name", String.class)))
            .writeTo(defaultTestPath);
    }

    @Test public void testConstructor9() throws IOException {
        new Workbook("test list map sheet Constructor9", author)
            .watch(Print::println)
            .setAutoSize(true)
            .addSheet(new ListMapSheet(createTestData(10)
                , WaterMark.of(author)
                , new Sheet.Column("ID", "id", int.class)
                , new Sheet.Column("NAME", "name", String.class)))
            .writeTo(defaultTestPath);
    }

    @Test public void testConstructor10() throws IOException {
        new Workbook("test list map sheet Constructor10", author)
            .watch(Print::println)
            .setAutoSize(true)
            .addSheet(new ListMapSheet("MAP"
                , createTestData(10)
                , WaterMark.of(author)
                , new Sheet.Column("ID", "id", int.class)
                , new Sheet.Column("NAME", "name", String.class)))
            .writeTo(defaultTestPath);
    }

    static List<Map<String, ?>> createTestData() {
        int size = random.nextInt(100) + 1;
        return createTestData(size);
    }

    static List<Map<String, ?>> createTestData(int size) {
        List<Map<String, ?>> list = new ArrayList<>(size);
        for (int i = 0; i < size; i++) {
            Map<String, Object> map = new HashMap<>();
            map.put("id", random.nextInt());
            map.put("name", getRandomString());
            list.add(map);
        }
        return list;
    }

    static List<Map<String, ?>> createAllTypeData() {
        int size = random.nextInt(100) + 1;
        return createAllTypeData(size);
    }

    static List<Map<String, ?>> createAllTypeData(int size) {
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
}

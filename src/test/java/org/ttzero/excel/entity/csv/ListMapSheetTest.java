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

package org.ttzero.excel.entity.csv;

import org.junit.Test;
import org.ttzero.excel.Print;
import org.ttzero.excel.entity.Column;
import org.ttzero.excel.entity.ListMapSheet;
import org.ttzero.excel.entity.Sheet;
import org.ttzero.excel.entity.Workbook;
import org.ttzero.excel.entity.WorkbookTest;

import java.io.IOException;
import java.math.BigDecimal;
import java.sql.Time;
import java.sql.Timestamp;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.util.Arrays;
import java.util.Collections;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

import static org.ttzero.excel.entity.ListMapSheetTest.createAllTypeData;
import static org.ttzero.excel.entity.ListMapSheetTest.createNullTestData;
import static org.ttzero.excel.entity.ListMapSheetTest.createTestData;

/**
 * @author guanquan.wang at 2019-04-28 19:16
 */
public class ListMapSheetTest extends WorkbookTest {

    @Test public void testWrite() throws IOException {
        new Workbook("test map")
            .watch(Print::println)
            .addSheet(createTestData())
            .addSheet(createTestData())
            .saveAsCSV()
            .writeTo(getOutputTestPath());
    }

    @Test public void testAllType() throws IOException {
        new Workbook("test all type map")
            .watch(Print::println)
            .addSheet(createAllTypeData())
            .saveAsCSV()
            .writeTo(getOutputTestPath());
    }

    @Test public void testHeaderColumn() throws IOException {
        new Workbook("test header column map")
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
            .saveAsCSV()
            .writeTo(getOutputTestPath());
    }

    @Test public void testHeaderColumnBox() throws IOException {
        new Workbook("test header column box type map")
            .watch(Print::println)
            .addSheet(createAllTypeData()
                , new Column("Character", "cv", Character.class)
                , new Column("Short", "sv", Short.class)
                , new Column("Integer", "nv", Integer.class)
                , new Column("Long", "lv", Long.class)
                , new Column("Float", "fv", Float.class)
                , new Column("Double", "dv", Double.class)
            )
            .saveAsCSV()
            .writeTo(getOutputTestPath());
    }

    @Test public void testConstructor1() throws IOException {
        new Workbook("test list map sheet Constructor1")
            .watch(Print::println)
            .addSheet(new ListMapSheet())
            .saveAsCSV()
            .writeTo(getOutputTestPath());
    }

    @Test public void testConstructor2() throws IOException {
        new Workbook("test list map sheet Constructor2")
            .watch(Print::println)
            .addSheet(new ListMapSheet("Map").setData(createTestData(10)))
            .saveAsCSV()
            .writeTo(getOutputTestPath());
    }

    @Test public void testConstructor3() throws IOException {
        new Workbook("test list map sheet Constructor3")
            .watch(Print::println)
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
            .saveAsCSV()
            .writeTo(getOutputTestPath());
    }

    @Test public void testConstructor5() throws IOException {
        new Workbook("test list map sheet Constructor5")
            .watch(Print::println)
            .addSheet(new ListMapSheet(createAllTypeData(10)))
            .saveAsCSV()
            .writeTo(getOutputTestPath());
    }

    @Test public void testConstructor6() throws IOException {
        new Workbook("test list map sheet Constructor6")
            .watch(Print::println)
            .addSheet(new ListMapSheet("Map", createAllTypeData(10)))
            .writeTo(getOutputTestPath());
    }

    @Test public void testConstructor7() throws IOException {
        new Workbook("test list map sheet Constructor7")
            .watch(Print::println)
            .addSheet(new ListMapSheet(createAllTypeData(10)
                , new Column("Character", "cv", Character.class)
                , new Column("Short", "sv", Short.class)
                , new Column("Integer", "nv", Integer.class)
                , new Column("Long", "lv", Long.class)
                , new Column("Float", "fv", Float.class)
                , new Column("Double", "dv", Double.class)))
            .saveAsCSV()
            .writeTo(getOutputTestPath());
    }

    @Test public void testConstructor8() throws IOException {
        new Workbook("test list map sheet Constructor8")
            .watch(Print::println)
            .addSheet(new ListMapSheet("MAP", createTestData(10)
                , new Column("ID", "id", int.class)
                , new Column("NAME", "name", String.class)))
            .saveAsCSV()
            .writeTo(getOutputTestPath());
    }

    @Test public void testConstructor9() throws IOException {
        new Workbook("test list map sheet Constructor9")
            .watch(Print::println)
            .addSheet(new ListMapSheet(createTestData(10)
                , new Column("ID", "id")
                , new Column("NAME", "name")))
            .saveAsCSV()
            .writeTo(getOutputTestPath());
    }

    @Test public void testConstructor10() throws IOException {
        new Workbook("test list map sheet Constructor10")
            .watch(Print::println)
            .addSheet(new ListMapSheet("MAP"
                , createTestData(10)
                , new Column("ID", "id", int.class)
                , new Column("NAME", "name", String.class)))
            .saveAsCSV()
            .writeTo(getOutputTestPath());
    }

    @Test public void testArray() throws IOException {
        Map<String, Object> data1 = new HashMap<>();
        data1.put("id", 1);
        data1.put("name", "abc");

        Map<String, Object> data2 = new HashMap<>();
        data2.put("id", 2);
        data2.put("name", "xyz");
        new Workbook("ListMapSheet array to csv")
            .watch(Print::println)
            .addSheet(new ListMapSheet().setData(Arrays.asList(data1, data2)))
            .saveAsCSV()
            .writeTo(getOutputTestPath());
    }

    @Test public void testSingleList() throws IOException {
        Map<String, Object> data = new HashMap<>();
        data.put("id", 1);
        data.put("name", "abc");

        new Workbook("ListMapSheet single list to csv")
            .watch(Print::println)
            .addSheet(new ListMapSheet().setData(Collections.singletonList(data)))
            .saveAsCSV()
            .writeTo(getOutputTestPath());
    }

    @Test public void testNullValue() throws IOException {
        new Workbook("test map null value")
            .watch(Print::println)
            .addSheet(createNullTestData(10))
            .saveAsCSV()
            .writeTo(getOutputTestPath());
    }
}

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

import org.junit.FixMethodOrder;
import org.junit.runners.MethodSorters;
import org.junit.Test;
import org.ttzero.excel.reader.ExcelReader;
import org.ttzero.excel.reader.Row;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;

/**
 * @author guanquan.wang at 2019-04-29 21:36
 */
@FixMethodOrder(MethodSorters.NAME_ASCENDING)
public class EmptySheetTest extends WorkbookTest {
    @Test public void testEmpty() throws IOException {
        String fileName = "test empty.xlsx";
        new Workbook()
            .addSheet(new ListSheet<>())
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            assertEquals(reader.sheet(0).rows().count(), 0L);
        }
    }

    @Test public void testEmptyWithHeader() throws IOException {
        String fileName = "test empty header.xlsx";
        new Workbook()
            .addSheet(new ListSheet<>("Empty"
                , new Column("ID", Integer.class)
                , new Column("NAME", String.class)
                , new Column("AGE", Integer.class)
            ))
            .writeTo(defaultTestPath.resolve(fileName));

        // Check header row
        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            org.ttzero.excel.reader.Sheet sheet = reader.sheet(0);
            assertEquals("Empty", sheet.getName());
            Iterator<Row> iter = sheet.rows().iterator();
            assertTrue(iter.hasNext());
            org.ttzero.excel.reader.Row row = iter.next();
            assertEquals("ID", row.getString(0));
            assertEquals("NAME", row.getString(1));
            assertEquals("AGE", row.getString(2));
        }
    }

    @Test public void testEmptyDataReader() throws IOException {
        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("test empty.xlsx"))) {
            long count = reader.sheets().flatMap(org.ttzero.excel.reader.Sheet::dataRows).count();
            assertEquals(count, 0L);
        }
    }

    @Test public void testEmptyWithHeaderDataReader() throws IOException {
        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("test empty header.xlsx"))) {
            long count = reader.sheets().flatMap(org.ttzero.excel.reader.Sheet::dataRows).count();
            assertEquals(count, 0L);
        }
    }

    @Test public void testEmptySheetSpecifyColumns() throws IOException {
        String fileName = "empty sheet specify columns.xlsx";
        new Workbook().setAutoSize(true)
            .addSheet(new ListSheet<ListObjectSheetTest.Item>(
                    new Column("id"), new Column("name")
                ).setData(new ArrayList<>())
            ).writeTo(defaultTestPath.resolve(fileName));

        // Check header row
        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            Iterator<Row> iter = reader.sheet(0).rows().iterator();
            assertTrue(iter.hasNext());
            org.ttzero.excel.reader.Row row = iter.next();
            assertEquals("id", row.getString(0));
            assertEquals("name", row.getString(1));
        }
    }

    @Test public void testEmptySheet() throws IOException {
        String fileName = "empty sheet not specify columns.xlsx";
        new Workbook().setAutoSize(true)
            .addSheet(new ListSheet<ListObjectSheetTest.Item>().setData(new ArrayList<>()))
            .writeTo(defaultTestPath.resolve(fileName));

        // Check header row
        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            assertEquals(reader.sheet(0).rows().count(), 0L);
        }
    }

    @Test public void testEmptySheetSubClassSpecified() throws IOException {
        String fileName = "empty sheet sub-class specified types.xlsx";
        new Workbook().setAutoSize(true)
            .addSheet(new ListSheet<ListObjectSheetTest.Item>() {
                @Override
                protected List<ListObjectSheetTest.Item> more() {
                    return super.more();
                }
            })
            .writeTo(defaultTestPath.resolve(fileName));

        // Check header row
        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            Iterator<Row> iter = reader.sheet(0).rows().iterator();
            assertTrue(iter.hasNext());
            org.ttzero.excel.reader.Row row = iter.next();
            assertEquals("id", row.getString(0));
            assertEquals("name", row.getString(1));
        }
    }

    @Test public void testEmptyMapList() throws IOException {
        String fileName = "empty map list sheet.xlsx";
        new Workbook().setAutoSize(true)
            .addSheet(new ListMapSheet<>("empty"))
            .writeTo(defaultTestPath.resolve(fileName));

        // Check header row
        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            assertEquals("empty", reader.sheet(0).getName());
            assertEquals(reader.sheet(0).rows().count(), 0L);
        }
    }

    @Test public void testEmptyMapListSpecifyHeaders() throws IOException {
        String fileName = "empty map list sheet specify headers.xlsx";
        new Workbook().setAutoSize(true)
            .addSheet(new ListMapSheet<>("empty", new Column("id"), new Column("name")))
            .writeTo(defaultTestPath.resolve(fileName));

        // Check header row
        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            assertEquals("empty", reader.sheet(0).getName());
            Iterator<Row> iter = reader.sheet(0).rows().iterator();
            assertTrue(iter.hasNext());
            org.ttzero.excel.reader.Row row = iter.next();
            assertEquals("id", row.getString(0));
            assertEquals("name", row.getString(1));
        }
    }

    @Test public void testSpecifyHeader() throws IOException {
        final String fileName = "empty sheet with simple header.xlsx";
        new Workbook().addSheet(new EmptySheet().setHeader("A", "B", "C", "D")).writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            Row row = reader.sheet(0).header(1).getHeader();
            assertEquals("A | B | C | D", row.toString());
        }
    }
}

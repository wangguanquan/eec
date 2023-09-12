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


package org.ttzero.excel.entity;

import org.junit.Test;
import org.ttzero.excel.reader.Cell;
import org.ttzero.excel.reader.ExcelReader;
import org.ttzero.excel.reader.HeaderRow;

import java.io.IOException;
import java.util.List;
import java.util.stream.Collectors;

/**
 * @author guanquan.wang at 2023-04-04 22:38
 */
public class ListObjectSheetTest2 extends WorkbookTest {
    @Test public void testSpecifyRowWrite() throws IOException {
        List<ListObjectSheetTest.Item> list = ListObjectSheetTest.Item.randomTestData();
        new Workbook().setAutoSize(true)
            .addSheet(new ListSheet<>(list).setStartRowIndex(4))
            .writeTo(defaultTestPath.resolve("test specify row 4 ListSheet.xlsx"));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("test specify row 4 ListSheet.xlsx"))) {
            List<ListObjectSheetTest.Item> readList = reader.sheet(0).header(4).rows().map(row -> row.to(ListObjectSheetTest.Item.class)).collect(Collectors.toList());
            assert list.size() == readList.size();
            for (int i = 0, len = list.size(); i < len; i++)
                assert list.get(i).equals(readList.get(i));
        }
    }

    @Test public void testSpecifyRowStayA1Write() throws IOException {
        List<ListObjectSheetTest.Item> list = ListObjectSheetTest.Item.randomTestData();
        new Workbook().setAutoSize(true)
            .addSheet(new ListSheet<>(list).setStartRowIndex(4, false))
            .writeTo(defaultTestPath.resolve("test specify row 4 stay A1 ListSheet.xlsx"));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("test specify row 4 stay A1 ListSheet.xlsx"))) {
            List<ListObjectSheetTest.Item> readList = reader.sheet(0).bind(ListObjectSheetTest.Item.class, 4).rows().map(row -> (ListObjectSheetTest.Item) row.get()).collect(Collectors.toList());
            assert list.size() == readList.size();
            for (int i = 0, len = list.size(); i < len; i++)
                assert list.get(i).equals(readList.get(i));
        }
    }

    @Test public void testSpecifyRowAndColWrite() throws IOException {
        List<ListObjectSheetTest.Item> list = ListObjectSheetTest.Item.randomTestData(10);
        new Workbook().setAutoSize(true)
            .addSheet(new ListSheet<ListObjectSheetTest.Item>("Item"
                , new Column("id", "id").setColIndex(3)
                , new Column("name", "name").setColIndex(4))
                .setData(list)
                .setStartRowIndex(4)
            ).writeTo(defaultTestPath.resolve("test specify row and cel ListSheet.xlsx"));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("test specify row and cel ListSheet.xlsx"))) {
            List<ListObjectSheetTest.Item> readList = reader.sheet(0).bind(ListObjectSheetTest.Item.class, 4).rows().map(row -> (ListObjectSheetTest.Item) row.get()).collect(Collectors.toList());
            assert list.size() == readList.size();
            for (int i = 0, len = list.size(); i < len; i++)
                assert list.get(i).equals(readList.get(i));
        }
    }

    @Test public void testSpecifyRowAndColStayA1Write() throws IOException {
        List<ListObjectSheetTest.Item> list = ListObjectSheetTest.Item.randomTestData(10);
        new Workbook().setAutoSize(true)
            .addSheet(new ListSheet<ListObjectSheetTest.Item>("Item"
                , new Column("id", "id").setColIndex(3)
                , new Column("name", "name").setColIndex(4))
                .setData(list)
                .setStartRowIndex(4, false)
            ).writeTo(defaultTestPath.resolve("test specify row and cel stay A1 ListSheet.xlsx"));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("test specify row and cel stay A1 ListSheet.xlsx"))) {
            List<ListObjectSheetTest.Item> readList = reader.sheet(0).bind(ListObjectSheetTest.Item.class, 4).rows().map(row -> (ListObjectSheetTest.Item) row.get()).collect(Collectors.toList());
            assert list.size() == readList.size();
            for (int i = 0, len = list.size(); i < len; i++)
                assert list.get(i).equals(readList.get(i));
        }
    }

    @Test public void testSpecifyRowIgnoreHeaderWrite() throws IOException {
        List<ListObjectSheetTest.Item> list = ListObjectSheetTest.Item.randomTestData();
        new Workbook().setAutoSize(true)
            .addSheet(new ListSheet<>(list).setStartRowIndex(4).ignoreHeader())
            .writeTo(defaultTestPath.resolve("test specify row 4 ignore header ListSheet.xlsx"));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("test specify row 4 ignore header ListSheet.xlsx"))) {
            List<ListObjectSheetTest.Item> readList = reader.sheet(0)
                .header(3)
                .bind(ListObjectSheetTest.Item.class, new HeaderRow().with(createHeaderRow()))
                .rows()
                .map(row -> (ListObjectSheetTest.Item) row.get())
                .collect(Collectors.toList());
            assert list.size() == readList.size();
            for (int i = 0, len = list.size(); i < len; i++)
                assert list.get(i).equals(readList.get(i));
        }
    }

    @Test public void testSpecifyRowStayA1IgnoreHeaderWrite() throws IOException {
        List<ListObjectSheetTest.Item> list = ListObjectSheetTest.Item.randomTestData();
        new Workbook().setAutoSize(true)
            .addSheet(new ListSheet<>(list).setStartRowIndex(4, false).ignoreHeader())
            .writeTo(defaultTestPath.resolve("test specify row 4 stay A1 ignore header ListSheet.xlsx"));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("test specify row 4 stay A1 ignore header ListSheet.xlsx"))) {
            List<ListObjectSheetTest.Item> readList = reader.sheet(0).bind(ListObjectSheetTest.Item.class, 3).rows().map(row -> (ListObjectSheetTest.Item) row.get()).collect(Collectors.toList());
            assert list.size() == readList.size();
            for (int i = 0, len = list.size(); i < len; i++)
                assert list.get(i).equals(readList.get(i));
        }
    }

    @Test public void testSpecifyRowAndColIgnoreHeaderWrite() throws IOException {
        List<ListObjectSheetTest.Item> list = ListObjectSheetTest.Item.randomTestData(10);
        new Workbook().setAutoSize(true)
            .addSheet(new ListSheet<ListObjectSheetTest.Item>("Item"
                , new Column("id", "id").setColIndex(3)
                , new Column("name", "name").setColIndex(4))
                .setData(list)
                .setStartRowIndex(4)
                .ignoreHeader()
            ).writeTo(defaultTestPath.resolve("test specify row and cel ignore header ListSheet.xlsx"));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("test specify row and cel ignore header ListSheet.xlsx"))) {
            List<ListObjectSheetTest.Item> readList = reader.sheet(0).bind(ListObjectSheetTest.Item.class, 3).rows().map(row -> (ListObjectSheetTest.Item) row.get()).collect(Collectors.toList());
            assert list.size() == readList.size();
            for (int i = 0, len = list.size(); i < len; i++)
                assert list.get(i).equals(readList.get(i));
        }
    }

    @Test public void testSpecifyRowAndColStayA1IgnoreHeaderWrite() throws IOException {
        List<ListObjectSheetTest.Item> list = ListObjectSheetTest.Item.randomTestData(10);
        new Workbook().setAutoSize(true)
            .addSheet(new ListSheet<ListObjectSheetTest.Item>("Item"
                , new Column("id", "id").setColIndex(3)
                , new Column("name", "name").setColIndex(4))
                .setData(list)
                .setStartRowIndex(4, false)
                .ignoreHeader()
            ).writeTo(defaultTestPath.resolve("test specify row and cel stay A1 ignore header ListSheet.xlsx"));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("test specify row and cel stay A1 ignore header ListSheet.xlsx"))) {
            List<ListObjectSheetTest.Item> readList = reader.sheet(0).bind(ListObjectSheetTest.Item.class, 3).rows().map(row -> (ListObjectSheetTest.Item) row.get()).collect(Collectors.toList());
            assert list.size() == readList.size();
            for (int i = 0, len = list.size(); i < len; i++)
                assert list.get(i).equals(readList.get(i));
        }
    }

    private static org.ttzero.excel.reader.Row createHeaderRow () {
        org.ttzero.excel.reader.Row headerRow = new org.ttzero.excel.reader.Row() {};
        Cell[] cells = new Cell[2];
        cells[0] = new Cell((short) 1).setSv("id");
        cells[1] = new Cell((short) 2).setSv("name");
        headerRow.setCells(cells);
        return headerRow;
    }
}

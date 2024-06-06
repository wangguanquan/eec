/*
 * Copyright (c) 2017-2022, guanquan.wang@yandex.com All Rights Reserved.
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
import org.ttzero.excel.annotation.FreezePanes;
import org.ttzero.excel.manager.Const;
import org.ttzero.excel.reader.ExcelReader;

import java.io.IOException;
import java.util.List;
import java.util.stream.Collectors;

import static org.junit.Assert.assertEquals;


/**
 * @author guanquan.wang at 2022-04-17 15:04
 */
public class FreezeSheetTest extends WorkbookTest {

    @Test public void testFreezeTopRow() throws IOException {
        String fileName = "Freeze Annotation Top Row.xlsx";
        List<ListObjectSheetTest.AllType> expectList = FreezeTopRow.randomTestData(FreezeTopRow::new);
        new Workbook()
                .addSheet(new ListSheet<>(expectList))
                .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            List<FreezeTopRow> list = reader.sheet(0).dataRows().map(row -> row.to(FreezeTopRow.class)).collect(Collectors.toList());
            assertEquals(expectList.size(), list.size());
            for (int i = 0, len = expectList.size(); i < len; i++) {
                FreezeTopRow expect = (FreezeTopRow) expectList.get(i), e = list.get(i);
                assertEquals(expect, e);
            }
        }
    }

    @Test public void testFreezeFirstColumn() throws IOException {
        String fileName = "Freeze Annotation First Column.xlsx";
        List<ListObjectSheetTest.AllType> expectList = FreezeTopRow.randomTestData(FreezeTopRow::new);
        new Workbook()
                .addSheet(new ListSheet<>(expectList))
                .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            List<FreezeTopRow> list = reader.sheet(0).dataRows().map(row -> row.to(FreezeTopRow.class)).collect(Collectors.toList());
            assertEquals(expectList.size(), list.size());
            for (int i = 0, len = expectList.size(); i < len; i++) {
                FreezeTopRow expect = (FreezeTopRow) expectList.get(i), e = list.get(i);
                assertEquals(expect, e);
            }
        }
    }


    @Test public void testFreezePanes11() throws IOException {
        String fileName = "Freeze Annotation Panes Row1 Column1.xlsx";
        List<ListObjectSheetTest.AllType> expectList = FreezePanesRow1Column1.randomTestData(FreezePanesRow1Column1::new);
        new Workbook(fileName)
                .addSheet(new ListSheet<>(expectList))
                .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            List<FreezePanesRow1Column1> list = reader.sheet(0).dataRows().map(row -> row.to(FreezePanesRow1Column1.class)).collect(Collectors.toList());
            assertEquals(expectList.size(), list.size());
            for (int i = 0, len = expectList.size(); i < len; i++) {
                FreezePanesRow1Column1 expect = (FreezePanesRow1Column1) expectList.get(i), e = list.get(i);
                assertEquals(expect, e);
            }
        }
    }

    @Test public void testFreezeByExtPropertyRow2() throws IOException {
        String fileName = "Freeze Top Row2.xlsx";
        List<ListObjectSheetTest.AllType> expectList = ListObjectSheetTest.AllType.randomTestData();
        new Workbook()
            .addSheet(new ListSheet<>(expectList)
                .putExtProp(Const.ExtendPropertyKey.FREEZE, Panes.row(2)))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            List<ListObjectSheetTest.AllType> list = reader.sheet(0).dataRows().map(row -> row.to(ListObjectSheetTest.AllType.class)).collect(Collectors.toList());
            assertEquals(expectList.size(), list.size());
            for (int i = 0, len = expectList.size(); i < len; i++) {
                ListObjectSheetTest.AllType expect = expectList.get(i), e = list.get(i);
                assertEquals(expect, e);
            }
        }
    }

    @Test public void testFreezeByExtPropertyColumn2() throws IOException {
        String fileName = "Freeze first Column2.xlsx";
        List<ListObjectSheetTest.AllType> expectList = ListObjectSheetTest.AllType.randomTestData();
        new Workbook()
            .addSheet(new ListSheet<>(expectList)
                .putExtProp(Const.ExtendPropertyKey.FREEZE, Panes.col(2)))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            List<ListObjectSheetTest.AllType> list = reader.sheet(0).dataRows().map(row -> row.to(ListObjectSheetTest.AllType.class)).collect(Collectors.toList());
            assertEquals(expectList.size(), list.size());
            for (int i = 0, len = expectList.size(); i < len; i++) {
                ListObjectSheetTest.AllType expect = expectList.get(i), e = list.get(i);
                assertEquals(expect, e);
            }
        }
    }

    @Test public void testFreezeByExtPropertyRow2Column2() throws IOException {
        String fileName = "Freeze Panes Row2 Column2.xlsx";
        List<ListObjectSheetTest.AllType> expectList = ListObjectSheetTest.AllType.randomTestData();
        new Workbook()
            .addSheet(new ListSheet<>(expectList)
                .putExtProp(Const.ExtendPropertyKey.FREEZE, Panes.of(2, 2)))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            List<ListObjectSheetTest.AllType> list = reader.sheet(0).dataRows().map(row -> row.to(ListObjectSheetTest.AllType.class)).collect(Collectors.toList());
            assertEquals(expectList.size(), list.size());
            for (int i = 0, len = expectList.size(); i < len; i++) {
                ListObjectSheetTest.AllType expect = expectList.get(i), e = list.get(i);
                assertEquals(expect, e);
            }
        }
    }

    @Test public void testFreezeByExtPropertyRow2Column2FromRow4() throws IOException {
        String fileName = "Freeze Panes Row2 Column2 From Row4.xlsx";
        List<ListObjectSheetTest.AllType> expectList = ListObjectSheetTest.AllType.randomTestData();
        new Workbook()
            .addSheet(new ListSheet<>(expectList).setStartRowIndex(4)
                .putExtProp(Const.ExtendPropertyKey.FREEZE, Panes.row(4)))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            List<ListObjectSheetTest.AllType> list = reader.sheet(0).dataRows().map(row -> row.to(ListObjectSheetTest.AllType.class)).collect(Collectors.toList());
            assertEquals(expectList.size(), list.size());
            for (int i = 0, len = expectList.size(); i < len; i++) {
                ListObjectSheetTest.AllType expect = expectList.get(i), e = list.get(i);
                assertEquals(expect, e);
            }
        }
    }

    @Test public void testFreezePans52() throws IOException {
        String fileName = "Freeze Annotation Panes Row5 Column2.xlsx";
        List<ListObjectSheetTest.AllType> expectList = FreezePanesRow5Column2.randomTestData(FreezePanesRow5Column2::new);
        new Workbook()
                .addSheet(new ListSheet<>(expectList))
                .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            List<FreezePanesRow5Column2> list = reader.sheet(0).dataRows().map(row -> row.to(FreezePanesRow5Column2.class)).collect(Collectors.toList());
            assertEquals(expectList.size(), list.size());
            for (int i = 0, len = expectList.size(); i < len; i++) {
                FreezePanesRow5Column2 expect = (FreezePanesRow5Column2) expectList.get(i), e = list.get(i);
                assertEquals(expect, e);
            }
        }
    }

    // Freeze the first row
    @FreezePanes(topRow = 1)
    public static class FreezeTopRow extends ListObjectSheetTest.AllType { }

    // Freeze the first column
    @FreezePanes(firstColumn = 1)
    public static class FreezeFirstColumn extends ListObjectSheetTest.AllType { }

    // Freeze the first row and first column
    @FreezePanes(topRow = 1, firstColumn = 1)
    public static class FreezePanesRow1Column1 extends ListObjectSheetTest.AllType { }

    // Freeze the 1-5 rows and 1-2 columns
    @FreezePanes(topRow = 5, firstColumn = 2)
    public static class FreezePanesRow5Column2 extends ListObjectSheetTest.AllType { }

}

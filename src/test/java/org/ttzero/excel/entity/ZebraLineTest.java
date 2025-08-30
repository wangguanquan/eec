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
import org.ttzero.excel.entity.e7.XMLZebraLineCellValueAndStyle;
import org.ttzero.excel.entity.style.Fill;
import org.ttzero.excel.entity.style.PatternType;
import org.ttzero.excel.entity.style.Styles;
import org.ttzero.excel.reader.ExcelReader;
import org.ttzero.excel.reader.Row;

import java.awt.Color;
import java.io.IOException;
import java.util.stream.Stream;

import static org.junit.Assert.assertTrue;


/**
 * @author guanquan.wang at 2023-03-03 11:03
 */
public class ZebraLineTest extends WorkbookTest {

    @Test public void testDefaultZebraLineOnWorkbook() throws IOException {
        final String fileName = "Default zebra line on workbook.xlsx";
        new Workbook().defaultZebraLine()
            .addSheet(new ListSheet<>(ListObjectSheetTest.Student.randomTestData()))
            .addSheet(new ListSheet<>(ListObjectSheetTest.Student.randomTestData()))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            assertZebraLineEquals(reader.sheet(0).header(1).rows(), PatternType.solid, new Color(233, 234, 236));
            assertZebraLineEquals(reader.sheet(1).header(1).rows(), PatternType.solid, new Color(233, 234, 236));
        }
    }

    @Test public void testDefaultZebraLineOnWorksheet() throws IOException {
        final String fileName = "Default zebra line on worksheet.xlsx";
        new Workbook()
            .addSheet(new ListSheet<>(ListObjectSheetTest.Student.randomTestData()))
            .addSheet(new ListSheet<>(ListObjectSheetTest.Student.randomTestData()).defaultZebraLine())
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            assertNonZebraLine(reader.sheet(0).header(1).rows());
            assertZebraLineEquals(reader.sheet(1).header(1).rows(), PatternType.solid, new Color(233, 234, 236));
        }
    }

    @Test public void testCustomZebraLineOnWorkbook() throws IOException {
        final String fileName = "Custom zebra line on workbook.xlsx";
        new Workbook().setZebraLine(new Fill(PatternType.lightHorizontal, Color.pink))
            .addSheet(new ListSheet<>(ListObjectSheetTest.Student.randomTestData()))
            .addSheet(new ListSheet<>(ListObjectSheetTest.Student.randomTestData()))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            assertZebraLineEquals(reader.sheet(0).header(1).rows(), PatternType.lightHorizontal, Color.pink);
            assertZebraLineEquals(reader.sheet(1).header(1).rows(), PatternType.lightHorizontal, Color.pink);
        }
    }

    @Test public void testCustomZebraLineOnWorksheet() throws IOException {
        final String fileName = "Custom zebra line on worksheet.xlsx";
        new Workbook()
            .addSheet(new ListSheet<>(ListObjectSheetTest.Student.randomTestData()))
            .addSheet(new ListSheet<>(ListObjectSheetTest.Student.randomTestData()).setZebraLine(new Fill(PatternType.lightHorizontal, Color.pink)))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            assertNonZebraLine(reader.sheet(0).header(1).rows());
            assertZebraLineEquals(reader.sheet(1).header(1).rows(), PatternType.lightHorizontal, Color.pink);
        }
    }

    @Test public void testCustomZebraLineOnWorksheet2() throws IOException {
        final String fileName = "Custom zebra line on worksheet2.xlsx";
        new Workbook()
            .addSheet(new ListSheet<>(ListObjectSheetTest.Student.randomTestData()))
            .addSheet(new ListSheet<>(ListObjectSheetTest.Student.randomTestData())
                .setCellValueAndStyle(new XMLZebraLineCellValueAndStyle(new Fill(PatternType.lightHorizontal, Color.pink))))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            assertNonZebraLine(reader.sheet(0).header(1).rows());
            assertZebraLineEquals(reader.sheet(1).header(1).rows(), PatternType.lightHorizontal, Color.pink);
        }
    }

    @Test public void testCustomZebraLineOnWorksheet3() throws IOException {
        final String fileName = "Custom zebra line on worksheet3.xlsx";
        new Workbook()
            .addSheet(new ListSheet<>(ListObjectSheetTest.Student.randomTestData()).setZebraLine(new Fill(PatternType.solid, Color.orange)))
            .addSheet(new ListSheet<>(ListObjectSheetTest.Student.randomTestData()).setZebraLine(new Fill(PatternType.lightHorizontal, Color.pink)))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            assertZebraLineEquals(reader.sheet(0).header(1).rows(), PatternType.solid, Color.orange);
            assertZebraLineEquals(reader.sheet(1).header(1).rows(), PatternType.lightHorizontal, Color.pink);
        }
    }

    @Test public void testCancelSpecifyWorksheet() throws IOException {
        final String fileName = "Cancel zebra line on worksheet.xlsx";
        new Workbook().defaultZebraLine()
            .addSheet(new ListSheet<>(ListObjectSheetTest.Student.randomTestData()).cancelZebraLine())
            .addSheet(new ListSheet<>(ListObjectSheetTest.Student.randomTestData()))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            assertNonZebraLine(reader.sheet(0).header(1).rows());
            assertZebraLineEquals(reader.sheet(1).header(1).rows(), PatternType.solid, new Color(233, 234, 236));
        }
    }

    static void assertNonZebraLine(Stream<Row> rows) {
        assertTrue(rows.allMatch(row -> {
            Styles styles = row.getStyles();
            int style = row.getCellStyle(0);
            Fill fill = styles.getFill(style);
            return fill == null || fill.getPatternType() == PatternType.none;
        }));
    }

    static void assertZebraLineEquals(Stream<Row> rows, PatternType patternType, Color color) {
        assertTrue(rows.allMatch(row -> {
            Styles styles = row.getStyles();
            // Skip header
            int rowNum = row.getRowNum();
            int style = row.getCellStyle(0);
            Fill fill = styles.getFill(style);
            return (rowNum & 1) == 0 ? fill == null || fill.getPatternType() == PatternType.none : fill != null && fill.getPatternType() == patternType && color.equals(fill.getFgColor());
        }));
    }
}

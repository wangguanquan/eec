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
import org.ttzero.excel.entity.e7.XMLCellValueAndStyle;
import org.ttzero.excel.entity.e7.XMLZebraLineCellValueAndStyle;
import org.ttzero.excel.entity.style.Fill;
import org.ttzero.excel.entity.style.PatternType;
import org.ttzero.excel.reader.ExcelReader;

import java.awt.Color;
import java.io.IOException;

import static org.ttzero.excel.entity.ZebraLineTest.assertNonZebraLine;
import static org.ttzero.excel.entity.ZebraLineTest.assertZebraLineEquals;

/**
 * @author guanquan.wang at 2023-02-24 17:26
 */
public class XMLCellValueAndStyleTest extends WorkbookTest {
    @Test public void testDefaultZebraLineWrite() throws IOException {
        final String fileName = "test default zebra-line.xlsx";
        new Workbook()
            .defaultZebraLine()
            .addSheet(new ListSheet<>(ListObjectSheetTest.Item.randomTestData()))
            .addSheet(new ListSheet<>(ListObjectSheetTest.Item.randomTestData()))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            assertZebraLineEquals(reader.sheet(0).header(1).rows(), PatternType.solid, new Color(233, 234, 236));
            assertZebraLineEquals(reader.sheet(1).header(1).rows(), PatternType.solid, new Color(233, 234, 236));
        }
    }

    @Test public void testCustomZebraLineWrite() throws IOException {
        final String fileName = "test orange zebra-line.xlsx";
        new Workbook()
            .setZebraLine(new Fill(PatternType.solid, Color.orange))
            .addSheet(new ListSheet<>(ListObjectSheetTest.Item.randomTestData()))
            .addSheet(new ListSheet<>(ListObjectSheetTest.Item.randomTestData()))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            assertZebraLineEquals(reader.sheet(0).header(1).rows(), PatternType.solid, Color.orange);
            assertZebraLineEquals(reader.sheet(1).header(1).rows(), PatternType.solid, Color.orange);
        }
    }

    @Test public void testCustomZebraLineCancelPartWrite() throws IOException {
        final String fileName = "test none origin zebra-line cancel part.xlsx";
        new Workbook()
            .setZebraLine(new Fill(PatternType.solid, Color.orange))
            .addSheet(new ListSheet<>(ListObjectSheetTest.Item.randomTestData()).cancelZebraLine())
            .addSheet(new ListSheet<>(ListObjectSheetTest.Item.randomTestData()))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            assertNonZebraLine(reader.sheet(0).header(1).rows());
            assertZebraLineEquals(reader.sheet(1).header(1).rows(), PatternType.solid, Color.orange);
        }
    }

    @Test public void testCustomZebraLineCancelAllWrite() throws IOException {
        final String fileName = "test none zebra-line cancel all.xlsx";
        new Workbook()
            .setZebraLine(new Fill(PatternType.solid, Color.orange))
            .addSheet(new ListSheet<>(ListObjectSheetTest.Item.randomTestData()).cancelZebraLine())
            .addSheet(new ListSheet<>(ListObjectSheetTest.Item.randomTestData()).cancelZebraLine())
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            assertNonZebraLine(reader.sheet(0).header(1).rows());
            assertNonZebraLine(reader.sheet(1).header(1).rows());
        }
    }

    @Test public void testCustom2ZebraLineWrite() throws IOException {
        final String fileName = "test origin none zebra-line.xlsx";
        new Workbook()
            .addSheet(new ListSheet<>(ListObjectSheetTest.Item.randomTestData()).setZebraLine(new Fill(PatternType.solid, Color.orange)))
            .addSheet(new ListSheet<>(ListObjectSheetTest.Item.randomTestData()))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            assertZebraLineEquals(reader.sheet(0).header(1).rows(), PatternType.solid, Color.orange);
            assertNonZebraLine(reader.sheet(1).header(1).rows());
        }
    }

    @Test public void testCustom3ZebraLineWrite() throws IOException {
        final String fileName = "test orange default zebra-line.xlsx";
        new Workbook()
            .addSheet(new ListSheet<>(ListObjectSheetTest.Item.randomTestData()).setZebraLine(new Fill(PatternType.solid, Color.orange)))
            .addSheet(new ListSheet<>(ListObjectSheetTest.Item.randomTestData()).defaultZebraLine())
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            assertZebraLineEquals(reader.sheet(0).header(1).rows(), PatternType.solid, Color.orange);
            assertZebraLineEquals(reader.sheet(1).header(1).rows(), PatternType.solid, new Color(233, 234, 236));
        }
    }

    @Test public void testCustom4ZebraLineWrite() throws IOException {
        final String fileName = "test cancel origin default zebra-line.xlsx";
        new Workbook()
            .cancelZebraLine()
            .addSheet(new ListSheet<>(ListObjectSheetTest.Item.randomTestData()).setZebraLine(new Fill(PatternType.solid, Color.orange)))
            .addSheet(new ListSheet<>(ListObjectSheetTest.Item.randomTestData()).defaultZebraLine())
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            assertZebraLineEquals(reader.sheet(0).header(1).rows(), PatternType.solid, Color.orange);
            assertZebraLineEquals(reader.sheet(1).header(1).rows(), PatternType.solid, new Color(233, 234, 236));
        }
    }

    @Test public void testCustomCellValueAndStyleWrite() throws IOException {
        final String fileName = "test orange ZebraLineCellValueAndStyle.xlsx";
        new Workbook()
            .addSheet(new ListSheet<>(ListObjectSheetTest.Item.randomTestData())
                .setCellValueAndStyle(new XMLZebraLineCellValueAndStyle(new Fill(PatternType.solid, Color.orange))))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            assertZebraLineEquals(reader.sheet(0).header(1).rows(), PatternType.solid, Color.orange);
        }
    }

    @Test public void testCustomCellValueAndStyle2Write() throws IOException {
        final String fileName = "test custom CellValueAndStyle.xlsx";
        new Workbook()
            .addSheet(new ListSheet<>(ListObjectSheetTest.Item.randomTestData()).setCellValueAndStyle(new XMLCellValueAndStyle()))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            assertNonZebraLine(reader.sheet(0).header(1).rows());
        }
    }
}

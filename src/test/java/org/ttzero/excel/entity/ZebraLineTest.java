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

import java.awt.Color;
import java.io.IOException;

/**
 * @author guanquan.wang at 2023-03-03 11:03
 */
public class ZebraLineTest extends WorkbookTest {

    @Test public void testDefaultZebraLineOnWorkbook() throws IOException {
        new Workbook("Default zebra line on workbook").defaultZebraLine()
            .addSheet(new ListSheet<>(ListObjectSheetTest.Student.randomTestData()))
            .addSheet(new ListSheet<>(ListObjectSheetTest.Student.randomTestData()))
            .writeTo(defaultTestPath);
    }

    @Test public void testDefaultZebraLineOnWorksheet() throws IOException {
        new Workbook("Default zebra line on worksheet")
            .addSheet(new ListSheet<>(ListObjectSheetTest.Student.randomTestData()))
            .addSheet(new ListSheet<>(ListObjectSheetTest.Student.randomTestData()).defaultZebraLine())
            .writeTo(defaultTestPath);
    }

    @Test public void testCustomZebraLineOnWorkbook() throws IOException {
        new Workbook("Custom zebra line on workbook").setZebraLine(new Fill(PatternType.lightHorizontal, Color.pink))
            .addSheet(new ListSheet<>(ListObjectSheetTest.Student.randomTestData()))
            .addSheet(new ListSheet<>(ListObjectSheetTest.Student.randomTestData()))
            .writeTo(defaultTestPath);
    }

    @Test public void testCustomZebraLineOnWorksheet() throws IOException {
        new Workbook("Custom zebra line on worksheet")
            .addSheet(new ListSheet<>(ListObjectSheetTest.Student.randomTestData()))
            .addSheet(new ListSheet<>(ListObjectSheetTest.Student.randomTestData()).setZebraLine(new Fill(PatternType.lightHorizontal, Color.pink)))
            .writeTo(defaultTestPath);
    }

    @Test public void testCustomZebraLineOnWorksheet2() throws IOException {
        new Workbook("Custom zebra line on worksheet2")
            .addSheet(new ListSheet<>(ListObjectSheetTest.Student.randomTestData()))
            .addSheet(new ListSheet<>(ListObjectSheetTest.Student.randomTestData())
                .setCellValueAndStyle(new XMLZebraLineCellValueAndStyle(new Fill(PatternType.lightHorizontal, Color.pink))))
            .writeTo(defaultTestPath);
    }

    @Test public void testCustomZebraLineOnWorksheet3() throws IOException {
        new Workbook("Custom zebra line on worksheet3")
            .addSheet(new ListSheet<>(ListObjectSheetTest.Student.randomTestData()).setZebraLine(new Fill(PatternType.solid, Color.orange)))
            .addSheet(new ListSheet<>(ListObjectSheetTest.Student.randomTestData()).setZebraLine(new Fill(PatternType.lightHorizontal, Color.pink)))
            .writeTo(defaultTestPath);
    }

    @Test public void testCancelSpecifyWorksheet() throws IOException {
        new Workbook("Cancel zebra line on worksheet").defaultZebraLine()
            .addSheet(new ListSheet<>(ListObjectSheetTest.Student.randomTestData()).cancelZebraLine())
            .addSheet(new ListSheet<>(ListObjectSheetTest.Student.randomTestData()))
            .writeTo(defaultTestPath);
    }
}

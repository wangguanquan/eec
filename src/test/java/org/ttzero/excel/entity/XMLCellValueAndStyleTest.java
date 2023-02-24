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

import java.awt.Color;
import java.io.IOException;

/**
 * @author guanquan.wang at 2023-02-24 17:26
 */
public class XMLCellValueAndStyleTest extends WorkbookTest {
    @Test public void testDefaultZebraLineWrite() throws IOException {
        new Workbook("test default zebra-line")
            .defaultZebraLine()
            .addSheet(new ListSheet<>(ListObjectSheetTest.Item.randomTestData()))
            .addSheet(new ListSheet<>(ListObjectSheetTest.Item.randomTestData()))
            .writeTo(defaultTestPath);
    }

    @Test public void testCustomZebraLineWrite() throws IOException {
        new Workbook("test orange zebra-line")
            .setZebraLine(new Fill(PatternType.solid, Color.orange))
            .addSheet(new ListSheet<>(ListObjectSheetTest.Item.randomTestData()))
            .addSheet(new ListSheet<>(ListObjectSheetTest.Item.randomTestData()))
            .writeTo(defaultTestPath);
    }

    @Test public void testCustomZebraLineCancelPartWrite() throws IOException {
        new Workbook("test none origin zebra-line cancel part")
            .setZebraLine(new Fill(PatternType.solid, Color.orange))
            .addSheet(new ListSheet<>(ListObjectSheetTest.Item.randomTestData()).cancelZebraLine())
            .addSheet(new ListSheet<>(ListObjectSheetTest.Item.randomTestData()))
            .writeTo(defaultTestPath);
    }

    @Test public void testCustomZebraLineCancelAllWrite() throws IOException {
        new Workbook("test none zebra-line cancel all")
            .setZebraLine(new Fill(PatternType.solid, Color.orange))
            .addSheet(new ListSheet<>(ListObjectSheetTest.Item.randomTestData()).cancelZebraLine())
            .addSheet(new ListSheet<>(ListObjectSheetTest.Item.randomTestData()).cancelZebraLine())
            .writeTo(defaultTestPath);
    }

    @Test public void testCustom2ZebraLineWrite() throws IOException {
        new Workbook("test origin none zebra-line")
            .addSheet(new ListSheet<>(ListObjectSheetTest.Item.randomTestData()).setZebraLine(new Fill(PatternType.solid, Color.orange)))
            .addSheet(new ListSheet<>(ListObjectSheetTest.Item.randomTestData()))
            .writeTo(defaultTestPath);
    }

    @Test public void testCustom3ZebraLineWrite() throws IOException {
        new Workbook("test orange default zebra-line")
            .addSheet(new ListSheet<>(ListObjectSheetTest.Item.randomTestData()).setZebraLine(new Fill(PatternType.solid, Color.orange)))
            .addSheet(new ListSheet<>(ListObjectSheetTest.Item.randomTestData()).defaultZebraLine())
            .writeTo(defaultTestPath);
    }

    @Test public void testCustom4ZebraLineWrite() throws IOException {
        new Workbook("test cancel origin default zebra-line")
            .cancelZebraLine()
            .addSheet(new ListSheet<>(ListObjectSheetTest.Item.randomTestData()).setZebraLine(new Fill(PatternType.solid, Color.orange)))
            .addSheet(new ListSheet<>(ListObjectSheetTest.Item.randomTestData()).defaultZebraLine())
            .writeTo(defaultTestPath);
    }

    @Test public void testCustomCellValueAndStyleWrite() throws IOException {
        new Workbook("test orange ZebraLineCellValueAndStyle")
            .addSheet(new ListSheet<>(ListObjectSheetTest.Item.randomTestData())
                .setCellValueAndStyle(new XMLZebraLineCellValueAndStyle(new Fill(PatternType.solid, Color.orange))))
            .writeTo(defaultTestPath);
    }

    @Test public void testCustomCellValueAndStyle2Write() throws IOException {
        new Workbook("test custom CellValueAndStyle")
            .addSheet(new ListSheet<>(ListObjectSheetTest.Item.randomTestData()).setCellValueAndStyle(new XMLCellValueAndStyle()))
            .writeTo(defaultTestPath);
    }
}

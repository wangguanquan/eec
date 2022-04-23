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
import org.ttzero.excel.Print;
import org.ttzero.excel.annotation.FreezePanes;
import org.ttzero.excel.manager.Const;

import java.io.IOException;


/**
 * @author guanquan.wang at 2022-04-17 15:04
 */
public class FreezeSheetTest extends WorkbookTest {

    @Test public void testFreezeTopRow() throws IOException {
        new Workbook("Freeze Annotation Top Row", author)
                .watch(Print::println)
                .addSheet(new ListSheet<>(FreezeTopRow.randomTestData(FreezeTopRow::new)))
                .writeTo(defaultTestPath);
    }

    @Test public void testFreezeFirstColumn() throws IOException {
        new Workbook("Freeze Annotation First Column", author)
                .watch(Print::println)
                .addSheet(new ListSheet<>(FreezeFirstColumn.randomTestData(FreezeFirstColumn::new)))
                .writeTo(defaultTestPath);
    }


    @Test public void testFreezePanes11() throws IOException {
        new Workbook("Freeze Annotation Panes Row1 Column1", author)
                .watch(Print::println)
                .addSheet(new ListSheet<>(FreezePanesRow1Column1.randomTestData(FreezePanesRow1Column1::new)))
                .writeTo(defaultTestPath);
    }


    @Test public void testFreezePans52() throws IOException {
        new Workbook("Freeze Annotation Panes Row5 Column2", author)
                .watch(Print::println)
                .addSheet(new ListSheet<>(FreezePanesRow5Column2.randomTestData(FreezePanesRow5Column2::new)))
                .writeTo(defaultTestPath);
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

    @Test public void testFreezeByExtPropertyRow2() throws IOException {
        new Workbook("Freeze Top Row2", author)
                .watch(Print::println)
                .addSheet(new ListSheet<>(ListObjectSheetTest.AllType.randomTestData())
                        .putExtProp(Const.WorksheetExtendProperty.FREEZE, Panes.row(2)))
                .writeTo(defaultTestPath);
    }

    @Test public void testFreezeByExtPropertyColumn2() throws IOException {
        new Workbook("Freeze first Column2", author)
                .watch(Print::println)
                .addSheet(new ListSheet<>(ListObjectSheetTest.AllType.randomTestData())
                        .putExtProp(Const.WorksheetExtendProperty.FREEZE, Panes.col(2)))
                .writeTo(defaultTestPath);
    }

    @Test public void testFreezeByExtPropertyRow2Column2() throws IOException {
        new Workbook("Freeze Panes Row2 Column2", author)
                .watch(Print::println)
                .addSheet(new ListSheet<>(ListObjectSheetTest.AllType.randomTestData())
                        .putExtProp(Const.WorksheetExtendProperty.FREEZE, Panes.of(2, 2)))
                .writeTo(defaultTestPath);
    }
}

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
import org.ttzero.excel.annotation.ExcelColumn;
import org.ttzero.excel.entity.Column;
import org.ttzero.excel.entity.CustomColIndexTest;
import org.ttzero.excel.entity.CustomizeDataSourceSheet;
import org.ttzero.excel.entity.ExcelWriteException;
import org.ttzero.excel.entity.ListSheet;
import org.ttzero.excel.entity.MultiHeaderColumnsTest;
import org.ttzero.excel.entity.TooManyColumnsException;
import org.ttzero.excel.entity.WaterMark;
import org.ttzero.excel.entity.Workbook;
import org.ttzero.excel.entity.WorkbookTest;
import org.ttzero.excel.entity.style.Fill;
import org.ttzero.excel.entity.style.PatternType;
import org.ttzero.excel.entity.ListObjectSheetTest.AllType;
import org.ttzero.excel.entity.ListObjectSheetTest.Item;
import org.ttzero.excel.entity.ListObjectSheetTest.Student;
import org.ttzero.excel.entity.ListObjectSheetTest.BoxAllType;
import org.ttzero.excel.entity.ListObjectSheetTest.ExtItem;
import org.ttzero.excel.entity.ListObjectSheetTest.NoColumnAnnotation;
import org.ttzero.excel.entity.ListObjectSheetTest.NoColumnAnnotation2;
import org.ttzero.excel.util.CSVUtil;

import java.awt.Color;
import java.io.IOException;

import java.util.Arrays;
import java.util.Collections;
import java.util.List;


import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertThrows;
import static org.ttzero.excel.entity.ListObjectSheetTest.sp;
import static org.ttzero.excel.entity.ListObjectSheetTest.conversion;

/**
 * @author guanquan.wang at 2019-04-28 19:17
 */
public class ListObjectSheetTest extends WorkbookTest {
    @Test public void testWrite() throws IOException {
        new Workbook()
            .addSheet(new ListSheet<>(Item.randomTestData()))
            .saveAsCSV()
            .writeTo(getOutputTestPath().resolve("test object.csv"));
    }

    @Test public void testAllTypeWrite() throws IOException {
        new Workbook()
            .addSheet(new ListSheet<>(AllType.randomTestData()))
            .saveAsCSV()
            .writeTo(getOutputTestPath().resolve("all type object.csv"));
    }

    @Test public void testAnnotation() throws IOException {
        new Workbook()
            .addSheet(new ListSheet<>(Student.randomTestData()))
            .saveAsCSV()
            .writeTo(getOutputTestPath().resolve("annotation object.csv"));
    }

    @Test public void testAnnotationAutoSize() throws IOException {
        new Workbook()
            .addSheet(new ListSheet<>(Student.randomTestData()))
            .saveAsCSV()
            .writeTo(getOutputTestPath().resolve("annotation object auto-size.csv"));
    }

    @Test public void testAutoSize() throws IOException {
        new Workbook()
            .addSheet(new ListSheet<>(AllType.randomTestData()))
            .saveAsCSV()
            .writeTo(getOutputTestPath().resolve("all type auto size.csv"));
    }

    @Test public void testIntConversion() throws IOException {
        new Workbook()
            .addSheet(new ListSheet<>(Student.randomTestData()
                , new Column("学号", "id")
                , new Column("姓名", "name")
                , new Column("成绩", "score", n -> (int) n < 60 ? "不合格" : n)
            ))
            .saveAsCSV()
            .writeTo(getOutputTestPath().resolve("test int conversion.csv"));
    }

    @Test public void testStyleConversion() throws IOException {
        new Workbook()
            .addSheet(new ListSheet<>(Student.randomTestData()
                , new Column("学号", "id")
                , new Column("姓名", "name")
                , new Column("成绩", "score")
                    .setStyleProcessor((o, style, sst) -> {
                        if ((int) o < 60) {
                            style = sst.modifyFill(style, new Fill(PatternType.solid, Color.orange));
                        }
                        return style;
                    })
            ))
            .saveAsCSV()
            .writeTo(getOutputTestPath().resolve("object style processor.csv"));
    }

    @Test public void testConvertAndStyleConversion() throws IOException {
        new Workbook()
            .addSheet(new ListSheet<>(Student.randomTestData()
                , new Column("学号", "id")
                , new Column("姓名", "name")
                , new Column("成绩", "score", n -> (int) n < 60 ? "不合格" : n)
                    .setStyleProcessor((o, style, sst) -> {
                        if ((int) o < 60) {
                            style = sst.modifyFill(style, new Fill(PatternType.solid, new Color(246, 209, 139)));
                        }
                        return style;
                    })
            ))
            .saveAsCSV()
            .writeTo(getOutputTestPath().resolve("object style and style processor.csv"));
    }

    @Test public void testCustomizeDataSource() throws IOException {
        new Workbook()
            .addSheet(new CustomizeDataSourceSheet())
            .saveAsCSV()
            .writeTo(getOutputTestPath().resolve("customize datasource.csv"));
    }

    @Test public void testBoxAllTypeWrite() throws IOException {
        new Workbook()
            .addSheet(new ListSheet<>(BoxAllType.randomTestData()))
            .saveAsCSV()
            .writeTo(getOutputTestPath().resolve("box all type object.csv"));
    }

    @Test public void testArray() throws IOException {
        new Workbook()
            .addSheet(new ListSheet<>()
                .setData(Arrays.asList(new Item(1, "abc"), new Item(2, "xyz"))))
            .saveAsCSV()
            .writeTo(getOutputTestPath().resolve("ListObjectSheet array to csv.csv"));
    }

    @Test public void testSingleList() throws IOException {
        new Workbook()
            .addSheet(new ListSheet<>()
                .setData(Collections.singletonList(new Item(1, "a b c"))))
            .saveAsCSV()
            .writeTo(getOutputTestPath().resolve("ListObject single list to csv.csv"));
    }

    @Test
    public void testStyleConversion1() throws IOException {
        new Workbook()
            .setWaterMark(WaterMark.of(author))
            .addSheet(new ListSheet<>("期末成绩", Student.randomTestData()
                    , new Column("学号", "id")
                    , new Column("姓名", "name")
                    , new Column("成绩", "score", conversion)
                    .setStyleProcessor(sp)
                )
            )
            .saveAsCSV()
            .writeTo(getOutputTestPath().resolve("object style processor1.csv"));
    }

    @Test public void testNullValue() throws IOException {
        new Workbook()
            .addSheet(new ListSheet<>("EXT-ITEM", ExtItem.randomTestData(10)))
            .saveAsCSV()
            .writeTo(getOutputTestPath().resolve("test null value.csv"));
    }

    @Test public void testFieldUnDeclare() throws IOException {
        new Workbook()
            .addSheet(new ListSheet<>("期末成绩", Student.randomTestData()
                    , new Column("学号", "id")
                    , new Column("姓名", "name")
                    , new Column("成绩", "sore") // un-declare field
                )
            )
            .saveAsCSV()
            .writeTo(getOutputTestPath().resolve("field un-declare.csv"));
    }

    @Test public void testResetMethod() throws IOException {
        new Workbook()
            .addSheet(new ListSheet<Student>("重写期末成绩", Collections.singletonList(new Student(9527, author, 0) {
                    @Override
                    public int getScore() {
                        return 100;
                    }
                }))
            )
            .saveAsCSV()
            .writeTo(getOutputTestPath().resolve("重写期末成绩.csv"));
    }

    @Test public void testMethodAnnotation() throws IOException {
        new Workbook()
            .addSheet(new ListSheet<Student>("重写方法注解", Collections.singletonList(new Student(9527, author, 0) {
                @Override
                @ExcelColumn("ID")
                public int getId() {
                    return super.getId();
                }

                @Override
                @ExcelColumn("SCORE")
                public int getScore() {
                    return 97;
                }
            }))
            )
            .saveAsCSV()
            .writeTo(getOutputTestPath().resolve("重写方法注解.csv"));
    }

    @Test public void testNoForceExport() throws IOException {
        new Workbook()
                .addSheet(new ListSheet<>(NoColumnAnnotation.randomTestData()))
                .saveAsCSV()
                .writeTo(getOutputTestPath().resolve("testNoForceExport.csv"));

        try (CSVUtil.Reader reader = CSVUtil.newReader(getOutputTestPath().resolve("testNoForceExport.csv"))) {
            assertEquals(reader.stream().count(), 0L);
        }
    }

    @Test public void testForceExportOnWorkbook() throws IOException {
        int lines = random.nextInt(100) + 3;
        new Workbook()
                .forceExport()
                .addSheet(new ListSheet<>(NoColumnAnnotation.randomTestData(lines)))
                .saveAsCSV()
                .writeTo(getOutputTestPath().resolve("testForceExportOnWorkbook.csv"));
        try (CSVUtil.Reader reader = CSVUtil.newReader(getOutputTestPath().resolve("testForceExportOnWorkbook.csv"))) {
            assertEquals(reader.stream().count(), lines + 1);
        }
    }

    @Test public void testForceExportOnWorkSheet() throws IOException {
        int lines = random.nextInt(100) + 3;
        new Workbook()
                .addSheet(new ListSheet<>(NoColumnAnnotation.randomTestData(lines)).forceExport())
                .saveAsCSV()
                .writeTo(getOutputTestPath().resolve("testForceExportOnWorkSheet.csv"));
        try (CSVUtil.Reader reader = CSVUtil.newReader(getOutputTestPath().resolve("testForceExportOnWorkSheet.csv"))) {
            assertEquals(reader.stream().count(), lines + 1);
        }
    }

    @Test public void testForceExportOnWorkbook2() throws IOException {
        int lines = random.nextInt(100) + 3, lines2 = random.nextInt(100) + 4;
        new Workbook()
                .forceExport()
                .addSheet(new ListSheet<>(NoColumnAnnotation.randomTestData(lines)))
                .addSheet(new ListSheet<>(NoColumnAnnotation2.randomTestData(lines2)))
                .saveAsCSV()
                .writeTo(getOutputTestPath().resolve("testForceExportOnWorkbook2.csv"));
    }

    @Test public void testForceExportOnWorkbook2Cancel1() throws IOException {
        int lines = random.nextInt(100) + 3, lines2 = random.nextInt(100) + 4;
        new Workbook()
                .forceExport()
                .addSheet(new ListSheet<>(NoColumnAnnotation.randomTestData(lines)).cancelForceExport())
                .addSheet(new ListSheet<>(NoColumnAnnotation2.randomTestData(lines2)))
                .saveAsCSV()
                .writeTo(getOutputTestPath().resolve("testForceExportOnWorkbook2Cancel1.csv"));
    }

    @Test public void testForceExportOnWorkbook2Cancel2() throws IOException {
        int lines = random.nextInt(100) + 3, lines2 = random.nextInt(100) + 4;
        new Workbook()
                .forceExport()
                .addSheet(new ListSheet<>(NoColumnAnnotation.randomTestData(lines)).cancelForceExport())
                .addSheet(new ListSheet<>(NoColumnAnnotation2.randomTestData(lines2)).cancelForceExport())
                .saveAsCSV()
                .writeTo(getOutputTestPath().resolve("testForceExportOnWorkbook2Cancel2.csv"));
    }

    @Test public void testOrderColumn() throws IOException {
        new Workbook()
                .addSheet(new ListSheet<>(CustomColIndexTest.OrderEntry.randomTestData()))
                .saveAsCSV()
                .writeTo(defaultTestPath.resolve("Order column.csv"));
    }

    @Test public void testSameOrderColumn() throws IOException {
        new Workbook()
                .addSheet(new ListSheet<>(CustomColIndexTest.SameOrderEntry.randomTestData()))
                .saveAsCSV()
                .writeTo(defaultTestPath.resolve("Same order column.csv"));
    }

    @Test public void testFractureOrderColumn() throws IOException {
        new Workbook()
                .addSheet(new ListSheet<>(CustomColIndexTest.FractureOrderEntry.randomTestData()))
                .saveAsCSV()
                .writeTo(defaultTestPath.resolve("Fracture order column.csv"));
    }

    @Test public void testLargeOrderColumn() throws IOException {
        new Workbook()
                .addSheet(new ListSheet<>(CustomColIndexTest.LargeOrderEntry.randomTestData()))
                .saveAsCSV()
                .writeTo(defaultTestPath.resolve("Large order column.csv"));
    }

    @Test public void testOverLargeOrderColumn() {
       assertThrows(TooManyColumnsException.class, () -> new Workbook(("Over Large order column"))
           .addSheet(new ListSheet<>(CustomColIndexTest.OverLargeOrderEntry.randomTestData()))
           .saveAsCSV()
           .writeTo(defaultTestPath));
    }

    @Test public void testRepeatAnnotations() throws IOException {
        List<MultiHeaderColumnsTest.RepeatableEntry> list = MultiHeaderColumnsTest.RepeatableEntry.randomTestData();
        new Workbook()
            .addSheet(new ListSheet<>(list))
            .saveAsCSV()
            .writeTo(defaultTestPath.resolve("Repeat Columns Annotation.csv"));
    }

    @Test public void testWriteWithBom() throws IOException {
        new Workbook()
            .addSheet(new ListSheet<>(Item.randomTestData()))
            .saveAsCSVWithBom()
            .writeTo(getOutputTestPath().resolve("test object with utf8 bom.csv"));
    }
}

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

import org.junit.BeforeClass;
import org.junit.Test;
import org.ttzero.excel.annotation.ExcelColumn;
import org.ttzero.excel.entity.Column;
import org.ttzero.excel.entity.CustomColIndexTest;
import org.ttzero.excel.entity.CustomizeDataSourceSheet;
import org.ttzero.excel.entity.ExcelWriteException;
import org.ttzero.excel.entity.ListSheet;
import org.ttzero.excel.entity.MultiHeaderColumnsTest;
import org.ttzero.excel.entity.TooManyColumnsException;
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

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.List;


import static org.ttzero.excel.entity.ListObjectSheetTest.sp;
import static org.ttzero.excel.entity.ListObjectSheetTest.conversion;

/**
 * @author guanquan.wang at 2019-04-28 19:17
 */
public class ListObjectSheetTest extends WorkbookTest {

    @BeforeClass
    public static void setUp() {
        // 删除之前的测试文件，防止重复执行不正确
        List<String> randomExcelNameList = new ArrayList<>();
        randomExcelNameList.add("testForceExportOnWorkSheet.csv");
        randomExcelNameList.add("testForceExportOnWorkbook.csv");
        randomExcelNameList.add("testForceExportOnWorkbook2Cancel1.xlsx");
        for (String fileName : randomExcelNameList) {
            Path resolve = defaultTestPath.resolve(fileName);
            boolean deleted = false;
            try {
                deleted = Files.deleteIfExists(resolve);
            } catch (IOException e) {
                e.printStackTrace();
            }
            if (deleted) {
                System.out.println(fileName + "File deleted successfully.");
            } else {
                System.out.println(fileName + "File deletion failed.");
            }
        }
    }

    @Test
    public void testWrite() throws IOException {
        new Workbook("test object")
            .addSheet(new ListSheet<>(Item.randomTestData()))
            .saveAsCSV()
            .writeTo(getOutputTestPath());
    }

    @Test public void testAllTypeWrite() throws IOException {
        new Workbook("all type object")
            .addSheet(new ListSheet<>(AllType.randomTestData()))
            .saveAsCSV()
            .writeTo(getOutputTestPath());
    }

    @Test public void testAnnotation() throws IOException {
        new Workbook("annotation object")
            .addSheet(new ListSheet<>(Student.randomTestData()))
            .saveAsCSV()
            .writeTo(getOutputTestPath());
    }

    @Test public void testAnnotationAutoSize() throws IOException {
        new Workbook("annotation object auto-size")
            .addSheet(new ListSheet<>(Student.randomTestData()))
            .saveAsCSV()
            .writeTo(getOutputTestPath());
    }

    @Test public void testAutoSize() throws IOException {
        new Workbook("all type auto size")
            .addSheet(new ListSheet<>(AllType.randomTestData()))
            .saveAsCSV()
            .writeTo(getOutputTestPath());
    }

    @Test public void testIntConversion() throws IOException {
        new Workbook("test int conversion")
            .addSheet(new ListSheet<>(Student.randomTestData()
                , new Column("学号", "id")
                , new Column("姓名", "name")
                , new Column("成绩", "score", n -> (int) n < 60 ? "不合格" : n)
            ))
            .saveAsCSV()
            .writeTo(getOutputTestPath());
    }

    @Test public void testStyleConversion() throws IOException {
        new Workbook("object style processor")
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
            .writeTo(getOutputTestPath());
    }

    @Test public void testConvertAndStyleConversion() throws IOException {
        new Workbook("object style and style processor")
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
            .writeTo(getOutputTestPath());
    }

    @Test public void testCustomizeDataSource() throws IOException {
        new Workbook("customize datasource")
            .addSheet(new CustomizeDataSourceSheet())
            .saveAsCSV()
            .writeTo(getOutputTestPath());
    }

    @Test public void testBoxAllTypeWrite() throws IOException {
        new Workbook("box all type object")
            .addSheet(new ListSheet<>(BoxAllType.randomTestData()))
            .saveAsCSV()
            .writeTo(getOutputTestPath());
    }

    @Test public void testArray() throws IOException {
        new Workbook("ListObjectSheet array to csv")
            .addSheet(new ListSheet<>()
                .setData(Arrays.asList(new Item(1, "abc"), new Item(2, "xyz"))))
            .saveAsCSV()
            .writeTo(getOutputTestPath());
    }

    @Test public void testSingleList() throws IOException {
        new Workbook("ListObject single list to csv")
            .addSheet(new ListSheet<>()
                .setData(Collections.singletonList(new Item(1, "a b c"))))
            .saveAsCSV()
            .writeTo(getOutputTestPath());
    }

    @Test
    public void testStyleConversion1() throws IOException {
        new Workbook("object style processor1", author)
            .addSheet(new ListSheet<>("期末成绩", Student.randomTestData()
                    , new Column("学号", "id")
                    , new Column("姓名", "name")
                    , new Column("成绩", "score", conversion)
                    .setStyleProcessor(sp)
                )
            )
            .saveAsCSV()
            .writeTo(getOutputTestPath());
    }

    @Test public void testNullValue() throws IOException {
        new Workbook("test null value")
            .addSheet(new ListSheet<>("EXT-ITEM", ExtItem.randomTestData(10)))
            .saveAsCSV()
            .writeTo(getOutputTestPath());
    }

    @Test public void testFieldUnDeclare() throws IOException {
        try {
            new Workbook("field un-declare")
                .addSheet(new ListSheet<>("期末成绩", Student.randomTestData()
                        , new Column("学号", "id")
                        , new Column("姓名", "name")
                        , new Column("成绩", "sore") // un-declare field
                    )
                )
                .saveAsCSV()
                .writeTo(getOutputTestPath());
        } catch (ExcelWriteException e) {
            assert true;
        }
    }

    @Test public void testResetMethod() throws IOException {
        new Workbook("重写期末成绩")
            .addSheet(new ListSheet<Student>("重写期末成绩", Collections.singletonList(new Student(9527, author, 0) {
                    @Override
                    public int getScore() {
                        return 100;
                    }
                }))
            )
            .saveAsCSV()
            .writeTo(getOutputTestPath());
    }

    @Test public void testMethodAnnotation() throws IOException {
        new Workbook("重写方法注解")
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
            .writeTo(getOutputTestPath());
    }

    @Test public void testNoForceExport() throws IOException {
        new Workbook("testNoForceExport")
                .addSheet(new ListSheet<>(NoColumnAnnotation.randomTestData()))
                .saveAsCSV()
                .writeTo(getOutputTestPath());

        try (CSVUtil.Reader reader = CSVUtil.newReader(getOutputTestPath().resolve("testNoForceExport.csv"))) {
            assert reader.stream().count() == 0L;
        }
    }

    @Test public void testForceExportOnWorkbook() throws IOException {
        int lines = random.nextInt(100) + 3;
        new Workbook("testForceExportOnWorkbook")
                .forceExport()
                .addSheet(new ListSheet<>(NoColumnAnnotation.randomTestData(lines)))
                .saveAsCSV()
                .writeTo(getOutputTestPath());
        try (CSVUtil.Reader reader = CSVUtil.newReader(getOutputTestPath().resolve("testForceExportOnWorkbook.csv"))) {
            assert reader.stream().count() == lines + 1;
        }
    }

    @Test public void testForceExportOnWorkSheet() throws IOException {
        int lines = random.nextInt(100) + 3;
        new Workbook("testForceExportOnWorkSheet")
                .addSheet(new ListSheet<>(NoColumnAnnotation.randomTestData(lines)).forceExport())
                .saveAsCSV()
                .writeTo(getOutputTestPath());
        try (CSVUtil.Reader reader = CSVUtil.newReader(getOutputTestPath().resolve("testForceExportOnWorkSheet.csv"))) {
            assert reader.stream().count() == lines + 1;
        }
    }

    @Test public void testForceExportOnWorkbook2() throws IOException {
        int lines = random.nextInt(100) + 3, lines2 = random.nextInt(100) + 4;
        new Workbook("testForceExportOnWorkbook2")
                .forceExport()
                .addSheet(new ListSheet<>(NoColumnAnnotation.randomTestData(lines)))
                .addSheet(new ListSheet<>(NoColumnAnnotation2.randomTestData(lines2)))
                .saveAsCSV()
                .writeTo(getOutputTestPath());
    }

    @Test public void testForceExportOnWorkbook2Cancel1() throws IOException {
        int lines = random.nextInt(100) + 3, lines2 = random.nextInt(100) + 4;
        new Workbook("testForceExportOnWorkbook2Cancel1")
                .forceExport()
                .addSheet(new ListSheet<>(NoColumnAnnotation.randomTestData(lines)).cancelForceExport())
                .addSheet(new ListSheet<>(NoColumnAnnotation2.randomTestData(lines2)))
                .saveAsCSV()
                .writeTo(getOutputTestPath());
    }

    @Test public void testForceExportOnWorkbook2Cancel2() throws IOException {
        int lines = random.nextInt(100) + 3, lines2 = random.nextInt(100) + 4;
        new Workbook("testForceExportOnWorkbook2Cancel2")
                .forceExport()
                .addSheet(new ListSheet<>(NoColumnAnnotation.randomTestData(lines)).cancelForceExport())
                .addSheet(new ListSheet<>(NoColumnAnnotation2.randomTestData(lines2)).cancelForceExport())
                .saveAsCSV()
                .writeTo(getOutputTestPath());
    }

    @Test public void testOrderColumn() throws IOException {
        new Workbook(("Order column"))
                .addSheet(new ListSheet<>(CustomColIndexTest.OrderEntry.randomTestData()))
                .saveAsCSV()
                .writeTo(defaultTestPath);
    }

    @Test public void testSameOrderColumn() throws IOException {
        new Workbook(("Same order column"))
                .addSheet(new ListSheet<>(CustomColIndexTest.SameOrderEntry.randomTestData()))
                .saveAsCSV()
                .writeTo(defaultTestPath);
    }

    @Test public void testFractureOrderColumn() throws IOException {
        new Workbook(("Fracture order column"))
                .addSheet(new ListSheet<>(CustomColIndexTest.FractureOrderEntry.randomTestData()))
                .saveAsCSV()
                .writeTo(defaultTestPath);
    }

    @Test public void testLargeOrderColumn() throws IOException {
        new Workbook(("Large order column"))
                .addSheet(new ListSheet<>(CustomColIndexTest.LargeOrderEntry.randomTestData()))
                .saveAsCSV()
                .writeTo(defaultTestPath);
    }

    @Test public void testOverLargeOrderColumn() throws IOException {
        try {
            new Workbook(("Over Large order column"))
                    .addSheet(new ListSheet<>(CustomColIndexTest.OverLargeOrderEntry.randomTestData()))
                    .saveAsCSV()
                    .writeTo(defaultTestPath);
        } catch (TooManyColumnsException e) {
            assert true;
        }
    }

    @Test public void testRepeatAnnotations() throws IOException {
        List<MultiHeaderColumnsTest.RepeatableEntry> list = MultiHeaderColumnsTest.RepeatableEntry.randomTestData();
        new Workbook("Repeat Columns Annotation")
            .addSheet(new ListSheet<>(list))
            .saveAsCSV()
            .writeTo(defaultTestPath);
    }
}

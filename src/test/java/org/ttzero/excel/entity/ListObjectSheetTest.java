/*
 * Copyright (c) 2017-2019, guanquan.wang@hotmail.com All Rights Reserved.
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
import org.ttzero.excel.annotation.ExcelColumn;
import org.ttzero.excel.annotation.IgnoreExport;
import org.ttzero.excel.entity.style.Fill;
import org.ttzero.excel.entity.style.Font;
import org.ttzero.excel.entity.style.PatternType;
import org.ttzero.excel.entity.style.Styles;
import org.ttzero.excel.manager.docProps.Core;
import org.ttzero.excel.processor.ConversionProcessor;
import org.ttzero.excel.processor.StyleProcessor;
import org.ttzero.excel.reader.AppInfo;
import org.ttzero.excel.reader.Dimension;
import org.ttzero.excel.reader.Drawings;
import org.ttzero.excel.reader.ExcelReader;
import org.ttzero.excel.reader.ExcelReaderTest;
import org.ttzero.excel.reader.HeaderRow;

import java.awt.Color;
import java.io.IOException;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.nio.file.Files;
import java.sql.Time;
import java.sql.Timestamp;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.Date;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Optional;
import java.util.function.Supplier;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertTrue;
import static org.ttzero.excel.reader.ExcelReaderTest.testResourceRoot;

/**
 * @author guanquan.wang at 2019-04-28 19:17
 */
public class ListObjectSheetTest extends WorkbookTest {

    @Test public void testWrite() throws IOException {
        String fileName = "test object.xlsx";
        List<Item> expectList = Item.randomTestData();
        new Workbook()
            .addSheet(new ListSheet<>(expectList))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            List<Item> list =  reader.sheet(0).dataRows().map(row -> row.to(Item.class)).collect(Collectors.toList());
            assertEquals(expectList.size(), list.size());
            for (int i = 0, len = expectList.size(); i < len; i++) {
                Item expect = expectList.get(i), e = list.get(i);
                assertEquals(expect, e);
            }
        }
    }

    @Test public void testAllTypeWrite() throws IOException {
        String fileName = "all type object.xlsx";
        List<AllType> expectList = AllType.randomTestData();
        new Workbook()
            .addSheet(new ListSheet<>(expectList))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            List<AllType> list =  reader.sheet(0).dataRows().map(row -> row.to(AllType.class)).collect(Collectors.toList());
            assertEquals(expectList.size(), list.size());
            for (int i = 0, len = expectList.size(); i < len; i++) {
                AllType expect = expectList.get(i), e = list.get(i);
                assertEquals(expect, e);
            }
        }
    }

    @Test public void testAnnotation() throws IOException {
        List<Student> expectList = Student.randomTestData();
        String fileName = "annotation object.xlsx";
        new Workbook()
            .addSheet(new ListSheet<>(expectList))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            List<Student> list =  reader.sheet(0).dataRows().map(row -> row.to(Student.class)).collect(Collectors.toList());
            assertEquals(expectList.size(), list.size());
            for (int i = 0, len = expectList.size(); i < len; i++) {
                Student expect = expectList.get(i), e = list.get(i);
                expect.id = 0; // ID not exported
                assertEquals(expect, e);
            }
        }
    }

    @Test public void testAnnotationAutoSize() throws IOException {
        List<Student> expectList = Student.randomTestData();
        String fileName = "annotation object auto-size.xlsx";
        new Workbook()
            .setAutoSize(true)
            .addSheet(new ListSheet<>(expectList))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            List<Student> list =  reader.sheet(0).dataRows().map(row -> row.to(Student.class)).collect(Collectors.toList());
            assertEquals(expectList.size(), list.size());
            for (int i = 0, len = expectList.size(); i < len; i++) {
                Student expect = expectList.get(i), e = list.get(i);
                expect.id = 0; // ID not exported
                assertEquals(expect, e);
            }
        }
    }

    @Test public void testStringWatermark() throws IOException {
        String fileName = "object string watermark.xlsx";
        List<Item> expectList = Item.randomTestData();
        new Workbook()
            .setWatermark(Watermark.of("SECRET"))
            .addSheet(new ListSheet<>(expectList))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            List<Item> list =  reader.sheet(0).dataRows().map(row -> row.to(Item.class)).collect(Collectors.toList());
            assertEquals(expectList.size(), list.size());
            for (int i = 0, len = expectList.size(); i < len; i++) {
                Item expect = expectList.get(i), e = list.get(i);
                assertEquals(expect, e);
            }

            List<Drawings.Picture> pictures = reader.sheet(0).listPictures();
            assertEquals(pictures.size(), 1);
            assertTrue(pictures.get(0).isBackground());
        }
    }

    @Test public void testLocalPicWatermark() throws IOException {
        String fileName = "object local pic watermark.xlsx";
        List<Item> expectList = Item.randomTestData();
        new Workbook()
            .setWatermark(Watermark.of(testResourceRoot().resolve("mark.png")))
            .addSheet(new ListSheet<>(expectList))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            List<Item> list =  reader.sheet(0).dataRows().map(row -> row.to(Item.class)).collect(Collectors.toList());
            assertEquals(expectList.size(), list.size());
            for (int i = 0, len = expectList.size(); i < len; i++) {
                Item expect = expectList.get(i), e = list.get(i);
                assertEquals(expect, e);
            }

            List<Drawings.Picture> pictures = reader.sheet(0).listPictures();
            assertEquals(pictures.size(), 1);
            Drawings.Picture pic = pictures.get(0);
            assertTrue(pic.isBackground());
            assertEquals(Files.size(pic.getLocalPath()), Files.size(testResourceRoot().resolve("mark.png")));
            assertEquals(crc32(pic.getLocalPath()), crc32(testResourceRoot().resolve("mark.png")));
        }
    }

    @Test public void testStreamWatermark() throws IOException {
        String fileName = "object input stream watermark.xlsx";
        List<Item> expectList = Item.randomTestData();
        new Workbook()
            .setWatermark(Watermark.of(getClass().getClassLoader().getResourceAsStream("mark.png")))
            .addSheet(new ListSheet<>(expectList))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            List<Item> list =  reader.sheet(0).dataRows().map(row -> row.to(Item.class)).collect(Collectors.toList());
            assertEquals(expectList.size(), list.size());
            for (int i = 0, len = expectList.size(); i < len; i++) {
                Item expect = expectList.get(i), e = list.get(i);
                assertEquals(expect, e);
            }

            List<Drawings.Picture> pictures = reader.sheet(0).listPictures();
            assertEquals(pictures.size(), 1);
            Drawings.Picture pic = pictures.get(0);
            assertTrue(pic.isBackground());
            assertEquals(Files.size(pic.getLocalPath()), Files.size(testResourceRoot().resolve("mark.png")));
            assertEquals(crc32(pic.getLocalPath()), crc32(testResourceRoot().resolve("mark.png")));
        }
    }

    @Test public void testAutoSize() throws IOException {
        String fileName = "all type auto size.xlsx";
        List<AllType> expectList = AllType.randomTestData();
        new Workbook()
            .setAutoSize(true)
            .addSheet(new ListSheet<>(expectList))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            List<AllType> list =  reader.sheet(0).dataRows().map(row -> row.to(AllType.class)).collect(Collectors.toList());
            assertEquals(expectList.size(), list.size());
            for (int i = 0, len = expectList.size(); i < len; i++) {
                AllType expect = expectList.get(i), e = list.get(i);
                assertEquals(expect, e);
            }
        }
    }

    @Test public void testIntConversion() throws IOException {
        String fileName = "test int conversion.xlsx";
        List<Student> expectList = Student.randomTestData();
        new Workbook()
            .addSheet(new ListSheet<>(expectList
                , new Column("学号", "id")
                , new Column("姓名", "name")
                , new Column("成绩", "score", n -> (int) n < 60 ? "不合格" : n)
            ))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            org.ttzero.excel.reader.Sheet sheet = reader.sheet(0).header(1);
            org.ttzero.excel.reader.HeaderRow header = (HeaderRow) sheet.getHeader();
            assertEquals("学号", header.get(0));
            assertEquals("姓名", header.get(1));
            assertEquals("成绩", header.get(2));


            Iterator<org.ttzero.excel.reader.Row> iter = sheet.iterator();
            for (Student expect : expectList) {
                assertTrue(iter.hasNext());
                Map<String, Object> e = iter.next().toMap();
                assertEquals(expect.getId(), Integer.parseInt(e.get("学号").toString()));
                assertEquals(expect.getName(), e.get("姓名").toString());
                if (expect.getScore() < 60) {
                    assertEquals("不合格", e.get("成绩"));
                } else {
                    assertEquals(expect.getScore(), Integer.parseInt(e.get("成绩").toString()));
                }
            }
        }
    }

    @Test public void testStyleConversion() throws IOException {
        String fileName = "object style processor.xlsx";
        List<Student> expectList = Student.randomTestData();
        new Workbook()
            .addSheet(new ListSheet<>(expectList
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
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            org.ttzero.excel.reader.Sheet sheet = reader.sheet(0).header(1).bind(Student.class);
            org.ttzero.excel.reader.HeaderRow header = (HeaderRow) sheet.getHeader();
            assertEquals("学号", header.get(0));
            assertEquals("姓名", header.get(1));
            assertEquals("成绩", header.get(2));

            Iterator<org.ttzero.excel.reader.Row> iter = sheet.iterator();
            for (Student expect : expectList) {
                assertTrue(iter.hasNext());
                org.ttzero.excel.reader.Row row = iter.next();
                Student e = row.to(Student.class);
                expect.id = 0; // ID not exported
                assertEquals(expect, e);

                Styles styles = row.getStyles();
                int style = row.getCellStyle(2);
                Fill fill = styles.getFill(style);
                if (expect.getScore() < 60) {
                    assertTrue(fill != null && fill.getPatternType() == PatternType.solid && fill.getFgColor().equals(Color.orange));
                } else {
                    assertTrue(fill == null || fill.getPatternType() == PatternType.none);
                }
            }
        }
    }

    @Test public void testConvertAndStyleConversion() throws IOException {
        String fileName = "object style and style processor.xlsx";
        List<Student> expectList = Student.randomTestData();
        new Workbook()
            .addSheet(new ListSheet<>(expectList
                , new Column("学号", "id")
                , new Column("姓名", "name")
                , new Column("成绩", "score", n -> (int) n < 60 ? "不合格" : n)
                    .setStyleProcessor((o, style, sst) -> {
                        if ((int)o < 60) {
                            style = sst.modifyFill(style, new Fill(PatternType.solid, new Color(246, 209, 139)));
                        }
                        return style;
                    })
            ))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            org.ttzero.excel.reader.Sheet sheet = reader.sheet(0).header(1);
            org.ttzero.excel.reader.HeaderRow header = (HeaderRow) sheet.getHeader();
            assertEquals("学号", header.get(0));
            assertEquals("姓名", header.get(1));
            assertEquals("成绩", header.get(2));

            Iterator<org.ttzero.excel.reader.Row> iter = sheet.iterator();
            for (Student expect : expectList) {
                assertTrue(iter.hasNext());
                org.ttzero.excel.reader.Row row = iter.next();
                Map<String, Object> e = row.toMap();
                assertEquals(expect.getId(), Integer.parseInt(e.get("学号").toString()));
                assertEquals(expect.getName(), e.get("姓名").toString());
                if (expect.getScore() < 60) {
                    assertEquals("不合格", e.get("成绩"));
                } else {
                    assertEquals(expect.getScore(), Integer.parseInt(e.get("成绩").toString()));
                }

                Styles styles = row.getStyles();
                int style = row.getCellStyle(2);
                Fill fill = styles.getFill(style);
                if (expect.getScore() < 60) {
                    assertTrue(fill != null && fill.getPatternType() == PatternType.solid && fill.getFgColor().equals(new Color(246, 209, 139)));
                } else {
                    assertTrue(fill == null || fill.getPatternType() == PatternType.none);
                }
            }
        }
    }

    @Test public void testCustomizeDataSource() throws IOException {
        String fileName = "customize datasource.xlsx";
        List<Student> expectList = new ArrayList<>();
        new Workbook()
            .addSheet(new CustomizeDataSourceSheet() {
                @Override
                public List<Student> more() {
                    List<Student> list = super.more();
                    if (list != null) expectList.addAll(list);
                    return list;
                }
            })
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            List<Student> list =  reader.sheet(0).dataRows().map(row -> row.to(Student.class)).collect(Collectors.toList());
            assertEquals(expectList.size(), list.size());
            for (int i = 0, len = expectList.size(); i < len; i++) {
                Student expect = expectList.get(i), e = list.get(i);
                assertEquals(expect.getName(), e.getName());
                assertEquals(expect.getScore(), e.getScore());
            }
        }
    }

    @Test public void testBoxAllTypeWrite() throws IOException {
        String fileName = "box all type object.xlsx";
        List<BoxAllType> expectList = BoxAllType.randomTestData();
        new Workbook()
            .addSheet(new ListSheet<>(expectList))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            List<BoxAllType> list =  reader.sheet(0).dataRows().map(row -> row.to(BoxAllType.class)).collect(Collectors.toList());
            assertEquals(expectList.size(), list.size());
            for (int i = 0, len = expectList.size(); i < len; i++) {
                BoxAllType expect = expectList.get(i), e = list.get(i);
                assertEquals(expect, e);
            }
        }
    }

    // -----AUTO SIZE

    @Test public void testBoxAllTypeAutoSizeWrite() throws IOException {
        String fileName = "auto-size box all type object.xlsx";
        List<BoxAllType> expectList = BoxAllType.randomTestData();
        new Workbook()
            .setAutoSize(true)
            .addSheet(new ListSheet<>(expectList))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            List<BoxAllType> list =  reader.sheet(0).dataRows().map(row -> row.to(BoxAllType.class)).collect(Collectors.toList());
            assertEquals(expectList.size(), list.size());
            for (int i = 0, len = expectList.size(); i < len; i++) {
                BoxAllType expect = expectList.get(i), e = list.get(i);
                assertEquals(expect, e);
            }
        }
    }

    @Test public void testCustomizeDataSourceAutoSize() throws IOException {
        String fileName = "auto-size customize datasource.xlsx";
        List<Student> expectList = new ArrayList<>();
        new Workbook()
            .setAutoSize(true)
            .addSheet(new CustomizeDataSourceSheet() {
                @Override
                public List<Student> more() {
                    List<Student> list = super.more();
                    if (list != null) expectList.addAll(list);
                    return list;
                }
            })
            .writeTo(defaultTestPath.resolve(fileName));

        assertFalse(expectList.isEmpty());

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            List<Student> list =  reader.sheet(0).dataRows().map(row -> row.to(Student.class)).collect(Collectors.toList());
            assertEquals(expectList.size(), list.size());
            for (int i = 0, len = expectList.size(); i < len; i++) {
                Student expect = expectList.get(i), e = list.get(i);
                assertEquals(expect.getName(), e.getName());
                assertEquals(expect.getScore(), e.getScore());
            }
        }
    }

    @Test public void testConstructor1() throws IOException {
        String fileName = "test list sheet Constructor1.xlsx";
        List<Item> expectList = Item.randomTestData();
        new Workbook()
            .setAutoSize(true)
            .addSheet(new ListSheet<>(expectList))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            List<Item> list =  reader.sheet(0).dataRows().map(row -> row.to(Item.class)).collect(Collectors.toList());
            assertEquals(expectList.size(), list.size());
            for (int i = 0, len = expectList.size(); i < len; i++) {
                Item expect = expectList.get(i), e = list.get(i);
                assertEquals(expect, e);
            }
        }
    }

    @Test public void testConstructor2() throws IOException {
        String fileName = "test list sheet Constructor2.xlsx";
        List<Item> expectList = Item.randomTestData(10);
        new Workbook()
            .setAutoSize(true)
            .addSheet(new ListSheet<Item>("Item").setData(expectList))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            assertEquals("Item", reader.sheet(0).getName());
            List<Item> list =  reader.sheet(0).dataRows().map(row -> row.to(Item.class)).collect(Collectors.toList());
            assertEquals(expectList.size(), list.size());
            for (int i = 0, len = expectList.size(); i < len; i++) {
                Item expect = expectList.get(i), e = list.get(i);
                assertEquals(expect, e);
            }
        }
    }

    @Test public void testConstructor3() throws IOException {
        String fileName = "test list sheet Constructor3.xlsx";
        List<Item> expectList = Item.randomTestData(10);
        new Workbook()
            .setAutoSize(true)
            .addSheet(new ListSheet<Item>("Item"
                , new Column("ID", "id")
                , new Column("NAME", "name"))
                .setData(expectList))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            assertEquals("Item", reader.sheet(0).getName());
            List<Item> list =  reader.sheet(0).headerColumnIgnoreCase().dataRows().map(row -> row.to(Item.class)).collect(Collectors.toList());
            assertEquals(expectList.size(), list.size());
            for (int i = 0, len = expectList.size(); i < len; i++) {
                Item expect = expectList.get(i), e = list.get(i);
                assertEquals(expect, e);
            }
        }
    }

    @Test public void testConstructor4() throws IOException {
        String fileName = "test list sheet Constructor4.xlsx";
        List<Item> expectList = Item.randomTestData(10);
        new Workbook()
            .setAutoSize(true)
            .addSheet(new ListSheet<Item>("Item"
                , new Column("ID", "id")
                , new Column("NAME", "name"))
                .setData(expectList).setWatermark(Watermark.of(author)))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            org.ttzero.excel.reader.Sheet sheet = reader.sheet(0);
            assertEquals("Item", sheet.getName());
            List<Item> list =  sheet.headerColumnIgnoreCase().dataRows().map(row -> row.to(Item.class)).collect(Collectors.toList());
            assertEquals(expectList.size(), list.size());
            for (int i = 0, len = expectList.size(); i < len; i++) {
                Item expect = expectList.get(i), e = list.get(i);
                assertEquals(expect, e);
            }

            List<Drawings.Picture> pictures = sheet.listPictures();
            assertEquals(pictures.size(), 1);
            Drawings.Picture pic = pictures.get(0);
            assertTrue(pic.isBackground());
        }
    }

    @Test public void testConstructor5() throws IOException {
        String fileName = "test list sheet Constructor5.xlsx";
        List<Item> expectList = Item.randomTestData(10);
        new Workbook()
            .setAutoSize(true)
            .addSheet(new ListSheet<>(expectList))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            org.ttzero.excel.reader.Sheet sheet = reader.sheet(0);
            List<Item> list =  sheet.dataRows().map(row -> row.to(Item.class)).collect(Collectors.toList());
            assertEquals(expectList.size(), list.size());
            for (int i = 0, len = expectList.size(); i < len; i++) {
                Item expect = expectList.get(i), e = list.get(i);
                assertEquals(expect, e);
            }
        }
    }

    @Test public void testConstructor6() throws IOException {
        String fileName = "test list sheet Constructor6.xlsx";
        List<Item> expectList = Item.randomTestData(10);
        new Workbook()
            .setAutoSize(true)
            .addSheet(new ListSheet<>("ITEM", expectList))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            org.ttzero.excel.reader.Sheet sheet = reader.sheet(0);
            assertEquals("ITEM", sheet.getName());
            List<Item> list =  sheet.dataRows().map(row -> row.to(Item.class)).collect(Collectors.toList());
            assertEquals(expectList.size(), list.size());
            for (int i = 0, len = expectList.size(); i < len; i++) {
                Item expect = expectList.get(i), e = list.get(i);
                assertEquals(expect, e);
            }
        }
    }

    @Test public void testConstructor7() throws IOException {
        String fileName = "test list sheet Constructor7.xlsx";
        List<Item> expectList = Item.randomTestData(10);
        new Workbook()
            .setAutoSize(true)
            .addSheet(new ListSheet<>(expectList
                , new Column("ID", "id")
                , new Column("NAME", "name")))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            org.ttzero.excel.reader.Sheet sheet = reader.sheet(0);
            List<Item> list =  sheet.headerColumnIgnoreCase().dataRows().map(row -> row.to(Item.class)).collect(Collectors.toList());
            assertEquals(expectList.size(), list.size());
            for (int i = 0, len = expectList.size(); i < len; i++) {
                Item expect = expectList.get(i), e = list.get(i);
                assertEquals(expect, e);
            }
        }
    }

    @Test public void testConstructor8() throws IOException {
        String fileName = "test list sheet Constructor8.xlsx";
        List<Item> expectList = Item.randomTestData(10);
        new Workbook()
            .setAutoSize(true)
            .addSheet(new ListSheet<>("ITEM", expectList
                , new Column("ID", "id")
                , new Column("NAME", "name")))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            org.ttzero.excel.reader.Sheet sheet = reader.sheet(0);
            assertEquals("ITEM", sheet.getName());
            List<Item> list =  sheet.headerColumnIgnoreCase().dataRows().map(row -> row.to(Item.class)).collect(Collectors.toList());
            assertEquals(expectList.size(), list.size());
            for (int i = 0, len = expectList.size(); i < len; i++) {
                Item expect = expectList.get(i), e = list.get(i);
                assertEquals(expect, e);
            }
        }
    }

    @Test public void testConstructor9() throws IOException {
        String fileName = "test list sheet Constructor9.xlsx";
        List<Item> expectList = Item.randomTestData(10);
        new Workbook()
            .setAutoSize(true)
            .addSheet(new ListSheet<>(expectList
                , new Column("ID", "id")
                , new Column("NAME", "name")).setWatermark(Watermark.of(author)))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            org.ttzero.excel.reader.Sheet sheet = reader.sheet(0);
            List<Item> list =  sheet.headerColumnIgnoreCase().dataRows().map(row -> row.to(Item.class)).collect(Collectors.toList());
            assertEquals(expectList.size(), list.size());
            for (int i = 0, len = expectList.size(); i < len; i++) {
                Item expect = expectList.get(i), e = list.get(i);
                assertEquals(expect, e);
            }
        }
    }

    @Test public void testConstructor10() throws IOException {
        String fileName = "test list sheet Constructor10.xlsx";
        List<Item> expectList = Item.randomTestData(10);
        new Workbook()
            .setAutoSize(true)
            .addSheet(new ListSheet<>("ITEM"
                , expectList
                , new Column("ID", "id")
                , new Column("NAME", "name")).setWatermark(Watermark.of(author)))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            org.ttzero.excel.reader.Sheet sheet = reader.sheet(0);
            assertEquals("ITEM", sheet.getName());
            List<Item> list =  sheet.headerColumnIgnoreCase().dataRows().map(row -> row.to(Item.class)).collect(Collectors.toList());
            assertEquals(expectList.size(), list.size());
            for (int i = 0, len = expectList.size(); i < len; i++) {
                Item expect = expectList.get(i), e = list.get(i);
                assertEquals(expect, e);
            }
            List<Drawings.Picture> pictures = sheet.listPictures();
            assertEquals(pictures.size(), 1);
            Drawings.Picture pic = pictures.get(0);
            assertTrue(pic.isBackground());
        }
    }

    @Test public void testArray() throws IOException {
        String fileName = "ListSheet Array as List.xlsx";
        List<Item> expectList = Arrays.asList(new Item(1, "abc"), new Item(2, "xyz"));
        new Workbook()
            .addSheet(new ListSheet<Item>()
                .setData(expectList))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            org.ttzero.excel.reader.Sheet sheet = reader.sheet(0);
            List<Item> list =  sheet.dataRows().map(row -> row.to(Item.class)).collect(Collectors.toList());
            assertEquals(expectList.size(), list.size());
            for (int i = 0, len = expectList.size(); i < len; i++) {
                Item expect = expectList.get(i), e = list.get(i);
                assertEquals(expect, e);
            }
        }
    }

    @Test public void testSingleList() throws IOException {
        String fileName = "ListSheet Single List.xlsx";
        List<Item> expectList = Collections.singletonList(new Item(1, "a b c"));
        new Workbook()
            .addSheet(new ListSheet<Item>()
                .setData(expectList))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            org.ttzero.excel.reader.Sheet sheet = reader.sheet(0);
            List<Item> list =  sheet.dataRows().map(row -> row.to(Item.class)).collect(Collectors.toList());
            assertEquals(expectList.size(), list.size());
            for (int i = 0, len = expectList.size(); i < len; i++) {
                Item expect = expectList.get(i), e = list.get(i);
                assertEquals(expect, e);
            }
        }
    }

    public static StyleProcessor<?> sp = (o, style, sst) -> {
        if ((int)o < 60) {
            style = sst.modifyFill(style, new Fill(PatternType.solid,Color.green, Color.blue));
        }
        return style;
    };

    // 定义一个int值转换lambda表达式，成绩低于60分显示"不合格"，其余显示正常分数
    public static ConversionProcessor conversion = n -> (int) n < 60 ? "不合格" : n;

    @Test public void testStyleConversion1() throws IOException {
        String fileName = "2021小五班期未考试成绩.xlsx";
        List<Student> expectList = Student.randomTestData();
        new Workbook()
            .addSheet(new ListSheet<>("期末成绩", expectList
                    , new Column("学号", "id", int.class)
                    , new Column("姓名", "name", String.class)
                    , new Column("成绩", "score", int.class, n -> (int) n < 60 ? "不合格" : n)
                ).setStyleProcessor((o, style, sst) ->
                    o.getScore() < 60 ? sst.modifyFill(style, new Fill(PatternType.solid, Color.orange)) : style)
            ).writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            org.ttzero.excel.reader.Sheet sheet = reader.sheet(0).header(1);
            assertEquals("期末成绩", sheet.getName());
            org.ttzero.excel.reader.HeaderRow header = (HeaderRow) sheet.getHeader();
            assertEquals("学号", header.get(0));
            assertEquals("姓名", header.get(1));
            assertEquals("成绩", header.get(2));

            Iterator<org.ttzero.excel.reader.Row> iter = sheet.iterator();
            for (Student expect : expectList) {
                assertTrue(iter.hasNext());
                org.ttzero.excel.reader.Row row = iter.next();
                assertEquals(expect.getId(), (long) row.getInt("学号"));
                assertEquals(expect.getName(), row.getString("姓名"));

                Styles styles = row.getStyles();
                if (expect.getScore() < 60) {
                    assertEquals("不合格", row.getString("成绩"));

                    int style0 = row.getCellStyle(0);
                    Fill fill0 = styles.getFill(style0);
                    assertTrue(fill0.getPatternType() == PatternType.solid && fill0.getFgColor().equals(Color.orange));
                    int style1 = row.getCellStyle(1);
                    Fill fill1 = styles.getFill(style1);
                    assertTrue(fill1.getPatternType() == PatternType.solid && fill1.getFgColor().equals(Color.orange));
                    int style2 = row.getCellStyle(2);
                    Fill fill2 = styles.getFill(style2);
                    assertTrue(fill2.getPatternType() == PatternType.solid && fill2.getFgColor().equals(Color.orange));
                } else {
                    assertEquals(expect.getScore(), (long) row.getInt("成绩"));

                    int style0 = row.getCellStyle(0);
                    Fill fill0 = styles.getFill(style0);
                    assertTrue(fill0 == null || fill0.getPatternType() == PatternType.none);
                    int style1 = row.getCellStyle(1);
                    Fill fill1 = styles.getFill(style1);
                    assertTrue(fill1 == null || fill1.getPatternType() == PatternType.none);
                    int style2 = row.getCellStyle(2);
                    Fill fill2 = styles.getFill(style2);
                    assertTrue(fill2 == null || fill2.getPatternType() == PatternType.none);
                }
            }
        }
    }

    @Test public void testNullValue() throws IOException {
        String fileName = "test null value.xlsx";
        List<Item> expectList = ExtItem.randomTestData(10);
        new Workbook()
            .setAutoSize(true)
            .addSheet(new ListSheet<>("EXT-ITEM", expectList))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            org.ttzero.excel.reader.Sheet sheet = reader.sheet(0);
            assertEquals("EXT-ITEM", sheet.getName());
            List<ExtItem> list =  sheet.dataRows().map(row -> row.to(ExtItem.class)).collect(Collectors.toList());
            assertEquals(expectList.size(), list.size());
            for (int i = 0, len = expectList.size(); i < len; i++) {
                ExtItem expect = (ExtItem) expectList.get(i), e = list.get(i);
                assertEquals(expect, e);
            }
        }
    }

    @Test public void testFieldUnDeclare() throws IOException {
        String fileName = "field un-declare.xlsx";
        List<Student> expectList = Student.randomTestData();
        new Workbook()
            .addSheet(new ListSheet<>("期末成绩", expectList
                    , new Column("学号", "id")
                    , new Column("姓名", "name")
                    , new Column("成绩", "score0") // un-declare field
                )
            )
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            org.ttzero.excel.reader.Sheet sheet = reader.sheet(0);
            assertEquals("期末成绩", sheet.getName());
            List<Map<String, Object>> list =  sheet.dataRows().map(org.ttzero.excel.reader.Row::toMap).collect(Collectors.toList());
            assertEquals(expectList.size(), list.size());
            for (int i = 0, len = expectList.size(); i < len; i++) {
                Student expect = expectList.get(i);
                Map<String, Object> e = list.get(i);
                assertEquals(expect.getId(), Integer.parseInt(e.get("学号").toString()));
                assertEquals(expect.getName(), e.get("姓名"));
                assertTrue(e.get("成绩") == null || e.get("成绩").equals(""));
            }
        }
    }

    @Test public void testResetMethod() throws IOException {
        String fileName = "重写期末成绩.xlsx";

        new Workbook()
            .addSheet(new ListSheet<>("重写期末成绩", Collections.singletonList(new Student(9527, author, 0) {
                    @Override
                    public int getScore() {
                        return 100;
                    }
                }))
            )
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            org.ttzero.excel.reader.Sheet sheet = reader.sheet(0);
            assertEquals("重写期末成绩", sheet.getName());
            List<Student> list =  sheet.dataRows().map(row -> row.to(Student.class)).collect(Collectors.toList());
            assertEquals(list.size(), 1);
            Student e = list.get(0);
            assertEquals(e.getId(), 0); // Ignore column
            assertEquals(author, e.getName());
            assertEquals(100, e.getScore());
        }
    }

    @Test public void testMethodAnnotation() throws IOException {
        String fileName = "重写方法注解.xlsx";
        new Workbook("重写方法注解", author)
            .addSheet(new ListSheet<>("重写方法注解", Collections.singletonList(new ExtStudent(9527, author, 0) {
                @Override
                @ExcelColumn("ID")
                public int getId() {
                    return super.getId();
                }

                @Override
                @ExcelColumn("成绩")
                public int getScore() {
                    return 97;
                }
            }))
            )
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            Optional<ExtStudent> opt = reader.sheets().flatMap(org.ttzero.excel.reader.Sheet::dataRows)
                .map(row -> row.too(ExtStudent.class)).findAny();
            assertTrue(opt.isPresent());
            ExtStudent student = opt.get();
            assertEquals(student.getId(), 9527);
            assertEquals(student.getScore(), 0); // The setter column name is 'score'
        }
    }

    @Test public void testIgnoreSupperMethod() throws IOException {
        final String fileName = "忽略父类属性.xlsx";
        new Workbook()
            .setWatermark(Watermark.of(author))
            .addSheet(new ListSheet<Student>("重写方法注解", Collections.singletonList(new ExtStudent(9527, author, 0))))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            Optional<ExtStudent> opt = reader.sheets().flatMap(org.ttzero.excel.reader.Sheet::dataRows)
                .map(row -> row.too(ExtStudent.class)).findAny();
            assertTrue(opt.isPresent());
            ExtStudent student = opt.get();
            assertEquals(student.getId(), 0);
            assertEquals(student.getScore(), 0);
        }
    }

    // Issue #93
    @Test public void testListSheet93() throws IOException {
        String fileName = "Issue#93 List Object.xlsx";
        List<Student> expectList = new ArrayList<>();
        new Workbook().addSheet(new ListSheet<Student>() {
            private int i;
            @Override
            protected List<Student> more() {
                List<Student> list = i++ < 10 ? Student.randomTestData(100) : null;
                if (list != null) expectList.addAll(list);
                return list;
            }
        }).writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            List<Student> list = reader.sheet(0).dataRows().map(row -> row.to(Student.class)).collect(Collectors.toList());
            assertEquals(expectList.size(), list.size());
            for (int i = 0, len = expectList.size(); i < len; i++) {
                Student expect = expectList.get(i), e = list.get(i);
                expect.id = 0; // ID not exported
                assertEquals(expect, e);
            }
        }
    }

    // Issue #95
    @Test public void testIssue95() throws IOException {
        String fileName = "Issue #95.xlsx";
        List<NotSharedObject> expectList = new ArrayList<>();
        new Workbook().addSheet(new ListSheet<NotSharedObject>() {
            private boolean c = true;
            @Override
            protected List<NotSharedObject> more() {
                if (!c) return null;
                c = false;
                List<NotSharedObject> list = new ArrayList<>();
                for (int i = 0; i < 10; i++) {
                    list.add(new NotSharedObject(getRandomString()));
                }
                expectList.addAll(list);
                return list;
            }
        }).writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            List<NotSharedObject> list = reader.sheet(0).dataRows().map(row -> row.to(NotSharedObject.class)).collect(Collectors.toList());
            assertEquals(expectList.size(), list.size());
            for (int i = 0, len = expectList.size(); i < len; i++) {
                NotSharedObject expect = expectList.get(i), e = list.get(i);
                assertEquals(expect, e);
            }
        }
    }

    @Test public void testSpecifyCore() throws IOException {
        final String fileName = "Specify Core.xlsx";
        Core core = new Core();
        core.setCreator("一名光荣的测试人员");
        core.setTitle("空白文件");
        core.setSubject("主题");
        core.setCategory("IT;木工");
        core.setDescription("为了艾尔");
        core.setKeywords("机枪兵;光头");
        core.setVersion("1.0");
//        core.setRevision("1.2");
        core.setLastModifiedBy("TTT");
        new Workbook().setCore(core)
            .addSheet(new ListSheet<>(Collections.singletonList(new NotSharedObject(getRandomString()))))
                .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            AppInfo info = reader.getAppInfo();
            assertEquals(core.getCreator(), info.getCreator());
            assertEquals(core.getTitle(), info.getTitle());
            assertEquals(core.getSubject(), info.getSubject());
            assertEquals(core.getCategory(), info.getCategory());
            assertEquals(core.getDescription(), info.getDescription());
            assertEquals(core.getKeywords(), info.getKeywords());
            assertEquals(core.getVersion(), info.getVersion());
            assertEquals(core.getLastModifiedBy(), info.getLastModifiedBy());
        }
    }

    @Test public void testLarge() throws IOException {
        final String fileName = "large07.xlsx";
        new Workbook().forceExport().addSheet(new ListSheet<ExcelReaderTest.LargeData>() {
            private int i = 0, n;

            @Override
            public List<ExcelReaderTest.LargeData> more() {
                if (n++ >= 10) return null;
                List<ExcelReaderTest.LargeData> list = new ArrayList<>();
                int size = i + 5000;
                for (; i < size; i++) {
                    ExcelReaderTest.LargeData largeData = new ExcelReaderTest.LargeData();
                    list.add(largeData);
                    largeData.setStr1("str1-" + i);
                    largeData.setStr2("str2-" + i);
                    largeData.setStr3("str3-" + i);
                    largeData.setStr4("str4-" + i);
                    largeData.setStr5("str5-" + i);
                    largeData.setStr6("str6-" + i);
                    largeData.setStr7("str7-" + i);
                    largeData.setStr8("str8-" + i);
                    largeData.setStr9("str9-" + i);
                    largeData.setStr10("str10-" + i);
                    largeData.setStr11("str11-" + i);
                    largeData.setStr12("str12-" + i);
                    largeData.setStr13("str13-" + i);
                    largeData.setStr14("str14-" + i);
                    largeData.setStr15("str15-" + i);
                    largeData.setStr16("str16-" + i);
                    largeData.setStr17("str17-" + i);
                    largeData.setStr18("str18-" + i);
                    largeData.setStr19("str19-" + i);
                    largeData.setStr20("str20-" + i);
                    largeData.setStr21("str21-" + i);
                    largeData.setStr22("str22-" + i);
                    largeData.setStr23("str23-" + i);
                    largeData.setStr24("str24-" + i);
                    largeData.setStr25("str25-" + i);
                }
                return list;
            }
        }).writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            org.ttzero.excel.reader.Sheet sheet = reader.sheet(0);
            assertEquals(Dimension.of("A1:Y50001"), sheet.getDimension());
            int i = 0;
            for (Iterator<org.ttzero.excel.reader.Row> iter = sheet.header(1).iterator(); iter.hasNext(); i++) {
                Map<String, Object> map = iter.next().toMap();
                assertEquals(map.get("str1"), "str1-" + i);
                assertEquals(map.get("str2"), "str2-" + i);
                assertEquals(map.get("str3"), "str3-" + i);
                assertEquals(map.get("str4"), "str4-" + i);
                assertEquals(map.get("str5"), "str5-" + i);
                assertEquals(map.get("str6"), "str6-" + i);
                assertEquals(map.get("str7"), "str7-" + i);
                assertEquals(map.get("str8"), "str8-" + i);
                assertEquals(map.get("str9"), "str9-" + i);
                assertEquals(map.get("str10"), "str10-" + i);
                assertEquals(map.get("str11"), "str11-" + i);
                assertEquals(map.get("str12"), "str12-" + i);
                assertEquals(map.get("str13"), "str13-" + i);
                assertEquals(map.get("str14"), "str14-" + i);
                assertEquals(map.get("str15"), "str15-" + i);
                assertEquals(map.get("str16"), "str16-" + i);
                assertEquals(map.get("str17"), "str17-" + i);
                assertEquals(map.get("str18"), "str18-" + i);
                assertEquals(map.get("str19"), "str19-" + i);
                assertEquals(map.get("str20"), "str20-" + i);
                assertEquals(map.get("str21"), "str21-" + i);
                assertEquals(map.get("str22"), "str22-" + i);
                assertEquals(map.get("str23"), "str23-" + i);
                assertEquals(map.get("str24"), "str24-" + i);
                assertEquals(map.get("str25"), "str25-" + i);
            }
        }
    }

    // #132
    @Test public void testEmptyList() throws IOException {
        String fileName = "ListObject empty list.xlsx";
        new Workbook().addSheet(new ListSheet<>(new ArrayList<>())).writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            assertEquals(reader.sheet(0).rows().count(), 0L);
        }
    }
    
    @Test public void testNoForceExport() throws IOException {
        String fileName = "testNoForceExport.xlsx";
        new Workbook().addSheet(new ListSheet<>(NoColumnAnnotation.randomTestData())).writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            assertEquals(Dimension.of("A1"), reader.sheet(0).getDimension());
            assertEquals(reader.sheet(0).rows().count(), 0L);
        }
    }
    
    @Test public void testForceExportOnWorkbook() throws IOException {
        String fileName = "testForceExportOnWorkbook.xlsx";
        int lines = random.nextInt(100) + 3;
        List<NoColumnAnnotation> expectList = NoColumnAnnotation.randomTestData(lines);
        new Workbook().forceExport().addSheet(new ListSheet<>(expectList)).writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            List<NoColumnAnnotation> list = reader.sheet(0).forceImport().dataRows().map(row -> row.to(NoColumnAnnotation.class)).collect(Collectors.toList());
            assertEquals(list.size(), lines);
            for (int i = 0; i < lines; i++) {
                NoColumnAnnotation expect = expectList.get(i), e = list.get(i);
                assertEquals(expect, e);
            }
        }
    }

    @Test public void testForceExportOnWorkSheet() throws IOException {
        String fileName = "testForceExportOnWorkSheet.xlsx";
        int lines = random.nextInt(100) + 3;
        List<NoColumnAnnotation> expectList = NoColumnAnnotation.randomTestData(lines);
        new Workbook().addSheet(new ListSheet<>(expectList).forceExport()).writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            List<NoColumnAnnotation> list = reader.sheet(0).forceImport().dataRows().map(row -> row.to(NoColumnAnnotation.class)).collect(Collectors.toList());
            assertEquals(list.size(), lines);
            for (int i = 0; i < lines; i++) {
                NoColumnAnnotation expect = expectList.get(i), e = list.get(i);
                assertEquals(expect, e);
            }
        }
    }

    @Test public void testForceExportOnWorkbook2() throws IOException {
        int lines = random.nextInt(100) + 3, lines2 = random.nextInt(100) + 4;
        String fileName = "testForceExportOnWorkbook2.xlsx";
        List<NoColumnAnnotation> expectList1 = NoColumnAnnotation.randomTestData(lines);
        List<NoColumnAnnotation2> expectList2 = NoColumnAnnotation2.randomTestData(lines2);
        new Workbook()
                .forceExport()
                .addSheet(new ListSheet<>(expectList1))
                .addSheet(new ListSheet<>(expectList2))
                .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            List<NoColumnAnnotation> list1 = reader.sheet(0).forceImport().dataRows().map(row -> row.to(NoColumnAnnotation.class)).collect(Collectors.toList());
            assertEquals(list1.size(), lines);
            for (int i = 0; i < lines; i++) {
                NoColumnAnnotation expect = expectList1.get(i), e = list1.get(i);
                assertEquals(expect, e);
            }

            List<NoColumnAnnotation2> list2 = reader.sheet(1).forceImport().dataRows().map(row -> row.to(NoColumnAnnotation2.class)).collect(Collectors.toList());
            assertEquals(list2.size(), lines2);
            for (int i = 0; i < lines2; i++) {
                NoColumnAnnotation2 expect = expectList2.get(i), e = list2.get(i);
                assertEquals(expect, e);
            }
        }
    }

    @Test public void testForceExportOnWorkbook2Cancel1() throws IOException {
        int lines = random.nextInt(100) + 3, lines2 = random.nextInt(100) + 4;
        String fileName = "testForceExportOnWorkbook2Cancel1.xlsx";
        List<NoColumnAnnotation> expectList1 = NoColumnAnnotation.randomTestData(lines);
        List<NoColumnAnnotation2> expectList2 = NoColumnAnnotation2.randomTestData(lines2);
        new Workbook()
                .forceExport()
                .addSheet(new ListSheet<>(expectList1).cancelForceExport())
                .addSheet(new ListSheet<>(expectList2))
                .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            assertEquals(reader.sheet(0).dataRows().count(), 0L);

            List<NoColumnAnnotation2> list2 = reader.sheet(1).forceImport().dataRows().map(row -> row.to(NoColumnAnnotation2.class)).collect(Collectors.toList());
            assertEquals(list2.size(), lines2);
            for (int i = 0; i < lines2; i++) {
                NoColumnAnnotation2 expect = expectList2.get(i), e = list2.get(i);
                assertEquals(expect, e);
            }
        }
    }

    @Test public void testForceExportOnWorkbook2Cancel2() throws IOException {
        int lines = random.nextInt(100) + 3, lines2 = random.nextInt(100) + 4;
        String fileName = "testForceExportOnWorkbook2Cancel2.xlsx";
        List<NoColumnAnnotation> expectList1 = NoColumnAnnotation.randomTestData(lines);
        List<NoColumnAnnotation2> expectList2 = NoColumnAnnotation2.randomTestData(lines2);
        new Workbook()
                .forceExport()
                .addSheet(new ListSheet<>(expectList1).cancelForceExport())
                .addSheet(new ListSheet<>(expectList2).cancelForceExport())
                .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            assertEquals(reader.sheet(0).dataRows().count(), 0L);
            assertEquals(reader.sheet(1).dataRows().count(), 0L);
        }
    }

    @Test public void testWrapText() throws IOException {
        String fileName = "WRAP TEXT.xlsx";
        List<Item> expectList;
        new Workbook()
            .addSheet(new ListSheet<Item>()
                .setData(expectList = Arrays.asList(new Item(1, "a b c\r\n1 2 3\r\n中文\t测试\r\nAAAAAA")
                    , new Item(2, "fdsafdsafdsafdsafdsafdsafdsafdsfadsafdsafdsafdsafdsafdsafds"))))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            List<Item> list = reader.sheet(0).dataRows().map(row -> row.to(Item.class)).collect(Collectors.toList());
            assertEquals(expectList.size(), list.size());
            for (int i = 0, len = expectList.size(); i < len; i++) {
                Item expect = expectList.get(i), e = list.get(i);
                assertEquals(expect, e);
            }
        }
    }

    @Test public void testClearHeadStyle() throws IOException {
        String fileName = "clear style.xlsx";
        Workbook workbook = new Workbook(fileName).addSheet(new ListSheet<>(Item.randomTestData()));

        Sheet sheet = workbook.getSheet(0);
        sheet.cancelZebraLine();  // Clear odd style
        int headStyle = sheet.defaultHeadStyle();
        sheet.setHeadStyle(Styles.clearFill(headStyle) & Styles.clearFont(headStyle));
        sheet.setHeadStyle(sheet.getHeadStyle() | workbook.getStyles().addFont(new Font("宋体", 11, Font.Style.BOLD, Color.BLACK)));
        workbook.writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            org.ttzero.excel.reader.HeaderRow header = (HeaderRow) reader.sheet(0).header(1).getHeader();
            Styles styles = header.getStyles();
            for (int i = header.getFirstColumnIndex(), limit = header.getLastColumnIndex(); i < limit; i++) {
                int styleIndex = header.getCellStyle(i);
                Fill fill = styles.getFill(styleIndex);
                assertTrue(fill == null || fill.getPatternType() == PatternType.none);
                Font font = styles.getFont(styleIndex);
                assertEquals(font.getSize(), 11);
                assertTrue(font.isBold());
                assertEquals(font.getColor(), Color.BLACK);
                assertEquals(font.getName(), "宋体");
            }
        }
    }

    @Test public void testBasicType() throws IOException {
        final String fileName = "Integer array.xlsx";
        List<Integer> list = new ArrayList<>(35);
        for (int i = 0; i < 35; i++) list.add(i);
        new Workbook()
            .addSheet(new ListSheet<>(list))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            Integer[] array = reader.sheets().flatMap(org.ttzero.excel.reader.Sheet::rows).map(row -> row.getInt(0)).toArray(Integer[]::new);
            assertEquals(array.length, list.size());
            for (int i = 0; i < array.length; i++) {
                assertEquals(array[i], list.get(i));
            }
        }
    }

    @Test public void testUnDisplayChar() throws Throwable {
        final String fileName = "UnDisplayChar.xlsx";
        List<Character> list = IntStream.range(0, 32).mapToObj(e -> (char)e).collect(Collectors.toList());
        new Workbook().addSheet(new ListSheet<>(list)).writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            List<Character> subList = reader.sheet(0).rows().map(row -> row.getChar(0)).collect(Collectors.toList());

            assertEquals(subList.size(), list.size());
            for (int i = 0; i < subList.size(); i++) {
                char c = subList.get(i);
                if (i == 9 || i == 10 || i == 13) {
                    assertEquals(list.get(i).charValue(), c);
                } else {
                    assertEquals(0xFFFD, c);
                }
            }
        }
    }

    @Test public void testEmojiChar() throws IOException {
        final String fileName = "Emoji char.xlsx";
        List<String> list = Arrays.asList("😂", "abc😍(●'◡'●)cz");
        new Workbook()
            .addSheet(new ListSheet<>(list).setColumns(new ListSheet.EntryColumn()))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            List<String> subList = reader.sheet(0).rows().map(row -> row.getString(0)).collect(Collectors.toList());

            assertEquals(subList.size(), list.size());

            for (int i = 0, len = subList.size(); i < len; i++) {
                assertEquals(subList.get(i), list.get(i));
            }
        }
    }

    @Test public void test264() throws IOException {
        String fileName = "Issue 264.xlsx";
        List<Item> expectList = Item.randomTestData(10);
        Column[] columns = {new Column("ID", "id"), new Column("NAME", "name")};
        new Workbook().addSheet(new ListSheet<>(expectList, columns))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            Iterator<org.ttzero.excel.reader.Row> iter = reader.sheet(0).iterator();
            assertTrue(iter.hasNext());
            org.ttzero.excel.reader.Row row = iter.next();
            assertEquals("ID", row.getString(0));
            assertEquals("NAME", row.getString(1));

            List<Item> list = reader.sheet(0).headerColumnIgnoreCase().dataRows().map(r -> r.to(Item.class)).collect(Collectors.toList());
            assertEquals(expectList.size(), list.size());
            for (int i = 0, len = expectList.size(); i < len; i++) {
                Item expect = expectList.get(i), e = list.get(i);
                assertEquals(expect, e);
            }
        }
    }

    @Test public void testNullInList() throws IOException {
        String fileName = "Null in list.xlsx";
        List<Item> expectList = Item.randomTestData(10);
        expectList.add(0, null);
        expectList.add(3, null);
        expectList.add(null);
        new Workbook().addSheet(new ListSheet<>(expectList)).writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            List<Item> list = reader.sheet(0).dataRows().map(row -> row.to(Item.class)).collect(Collectors.toList());
            assertEquals(list.size(), expectList.size() - 3);
            for (int i = 0, j = 0, len = expectList.size(); i < len; i++) {
                Item expect = expectList.get(i);
                if (expect == null) continue;
                Item e = list.get(j++);
                assertEquals(expect, e);
            }
        }
    }

    public static class Item {
        @ExcelColumn
        private int id;
        @ExcelColumn(wrapText = true)
        private String name;
        public Item() { }
        public Item(int id, String name) {
            this.id = id;
            this.name = name;
        }

        public void setId(int id) {
            this.id = id;
        }

        public void setName(String name) {
            this.name = name;
        }

        public int getId() {
            return id;
        }

        public String getName() {
            return name;
        }
        public static List<Item> randomTestData(int n) {
            return randomTestData(n, () -> new Item(random.nextInt(100), getRandomString()));
        }

        @Override
        public int hashCode() {
            return id ^ (name != null ? name.hashCode() : 0);
        }

        @Override
        public boolean equals(Object obj) {
            if (obj instanceof Item) {
                Item other = (Item) obj;
                return id == other.id && (name != null && name.equals(other.name) || name == null && other.name == null);
            }
            return false;
        }

        public static List<Item> randomTestData(int n, Supplier<Item> supplier) {
            List<Item> list = new ArrayList<>(n);
            for (int i = 0; i < n; i++) {
                list.add(supplier.get());
            }
            return list;
        }

        public static List<Item> randomTestData() {
            int n = random.nextInt(100) + 1;
            return randomTestData(n, () -> new Item(random.nextInt(100), getRandomString()));
        }
    }

    public static class AllType {
        @ExcelColumn
        private boolean bv;
        @ExcelColumn
        private char cv;
        @ExcelColumn
        private short sv;
        @ExcelColumn
        private int nv;
        @ExcelColumn
        private long lv;
        @ExcelColumn
        private float fv;
        @ExcelColumn
        private double dv;
        @ExcelColumn
        private String s;
        @ExcelColumn
        private BigDecimal mv;
        @ExcelColumn
        private Date av;
        @ExcelColumn
        private Timestamp iv;
        @ExcelColumn
        private Time tv;
        @ExcelColumn
        private LocalDate ldv;
        @ExcelColumn
        private LocalDateTime ldtv;
        @ExcelColumn
        private LocalTime ltv;

        public static List<AllType> randomTestData(int size, Supplier<AllType> sup) {
            List<AllType> list = new ArrayList<>(size);
            for (int i = 0; i < size; i++) {
                AllType o = sup.get();
                o.bv = random.nextInt(10) == 5;
                o.cv = charArray[random.nextInt(charArray.length)];
                o.sv = (short) (random.nextInt() & 0xFFFF);
                o.nv = random.nextInt();
                o.lv = random.nextLong();
                o.fv = random.nextFloat();
                o.dv = random.nextDouble();
                o.s = getRandomString();
                o.mv = BigDecimal.valueOf(random.nextDouble());
                o.av = new Date();
                o.iv = new Timestamp(System.currentTimeMillis() - random.nextInt(9999999));
                o.tv = new Time(random.nextLong());
                o.ldv = LocalDate.now();
                o.ldtv = LocalDateTime.now();
                o.ltv = LocalTime.now();
                list.add(o);
            }
            return list;
        }

        public static List<AllType> randomTestData() {
            return randomTestData(AllType::new);
        }

        public static List<AllType> randomTestData(Supplier<AllType> sup) {
            int size = random.nextInt(100) + 1;
            return randomTestData(size, sup);
        }

        public boolean isBv() {
            return bv;
        }

        public char getCv() {
            return cv;
        }

        public short getSv() {
            return sv;
        }

        public int getNv() {
            return nv;
        }

        public long getLv() {
            return lv;
        }

        public float getFv() {
            return fv;
        }

        public double getDv() {
            return dv;
        }

        public String getS() {
            return s;
        }

        public BigDecimal getMv() {
            return mv;
        }

        public Date getAv() {
            return av;
        }

        public Timestamp getIv() {
            return iv;
        }

        public Time getTv() {
            return tv;
        }

        public LocalDate getLdv() {
            return ldv;
        }

        public LocalDateTime getLdtv() {
            return ldtv;
        }

        public LocalTime getLtv() {
            return ltv;
        }

        public void setBv(boolean bv) {
            this.bv = bv;
        }

        public void setCv(char cv) {
            this.cv = cv;
        }

        public void setSv(short sv) {
            this.sv = sv;
        }

        public void setNv(int nv) {
            this.nv = nv;
        }

        public void setLv(long lv) {
            this.lv = lv;
        }

        public void setFv(float fv) {
            this.fv = fv;
        }

        public void setDv(double dv) {
            this.dv = dv;
        }

        public void setS(String s) {
            this.s = s;
        }

        public void setMv(BigDecimal mv) {
            this.mv = mv;
        }

        public void setAv(Date av) {
            this.av = av;
        }

        public void setIv(Timestamp iv) {
            this.iv = iv;
        }

        public void setTv(Time tv) {
            this.tv = tv;
        }

        public void setLdv(LocalDate ldv) {
            this.ldv = ldv;
        }

        public void setLdtv(LocalDateTime ldtv) {
            this.ldtv = ldtv;
        }

        public void setLtv(LocalTime ltv) {
            this.ltv = ltv;
        }

        @Override
        public boolean equals(Object o) {
            if (this == o) return true;
            if (o == null || getClass() != o.getClass()) return false;
            AllType allType = (AllType) o;
            return bv == allType.bv &&
                cv == allType.cv &&
                sv == allType.sv &&
                nv == allType.nv &&
                lv == allType.lv &&
                Float.compare(allType.fv, fv) == 0 &&
                Double.compare(allType.dv, dv) == 0 &&
                Objects.equals(s, allType.s) &&
                Objects.equals(mv.setScale(4, RoundingMode.HALF_DOWN), allType.mv.setScale(4, RoundingMode.HALF_DOWN)) &&
                av.getTime() / 1000 == allType.av.getTime() / 1000 &&
                iv.getTime() / 1000 == allType.iv.getTime() / 1000 &&
                String.valueOf(tv).equals(String.valueOf(allType.tv)) &&
                Objects.equals(ldv, allType.ldv) &&
                Timestamp.valueOf(ldtv).getTime() / 1000 == Timestamp.valueOf(allType.ldtv).getTime() / 1000 &&
                String.valueOf(Time.valueOf(ltv)).equals(String.valueOf(Time.valueOf(allType.ltv)));
        }

        @Override
        public int hashCode() {
            return Objects.hash(bv, cv, sv, nv, lv, fv, dv, s, mv, av.getTime() / 1000, iv.getTime() / 1000, String.valueOf(tv), ldv, Timestamp.valueOf(ldtv).getTime() / 1000, String.valueOf(Time.valueOf(ltv)));
        }

        @Override
        public String toString() {
            return "" + bv + '|' + cv + '|' + sv + '|' + nv + '|' + lv
                + '|' + fv + '|' + dv + '|' + s + '|' + mv + '|' + av
                + '|' + tv + '|' + ldv + '|' + ldtv + '|' + ltv;
        }
    }

    public static class Student {
        @IgnoreExport
        private int id;
        @ExcelColumn("姓名")
        private String name;
        @ExcelColumn("成绩")
        private int score;

        public Student() { }

        protected Student(int id, String name, int score) {
            this.id = id;
            this.name = name;
            this.score = score;
        }

        public int getId() {
            return id;
        }

        public void setId(int id) {
            this.id = id;
        }

        public String getName() {
            return name;
        }

        public void setName(String name) {
            this.name = name;
        }

        public int getScore() {
            return score;
        }

        public void setScore(int score) {
            this.score = score;
        }

        public static List<Student> randomTestData(int pageNo, int limit) {
            List<Student> list = new ArrayList<>(limit);
            for (int i = pageNo * limit, n = i + limit, k; i < n; i++) {
                Student e = new Student(random.nextInt(100), (k = random.nextInt(10)) < 3 ? String.valueOf((char) ('a' + k)) : getRandomString(), random.nextInt(50) + 50);
                list.add(e);
            }
            return list;
        }

        public static List<Student> randomTestData(int n) {
            return randomTestData(0, n);
        }

        public static List<Student> randomTestData() {
            int n = random.nextInt(100) + 1;
            return randomTestData(n);
        }

        @Override
        public int hashCode() {
            return (getId() << 16 | getScore()) ^ getName().hashCode();
        }

        @Override
        public boolean equals(Object obj) {
            if (obj instanceof Student) {
                Student other = (Student) obj;
                return getId() == other.getId() && getScore() == other.getScore() && getName().equals(other.getName());
            }
            return false;
        }

        @Override
        @ExcelColumn
        public String toString() {
            return "id: " + getId() + ", name: " + getName() + ", score: " + getScore();
        }
    }

    public static class BoxAllType {
        @ExcelColumn
        private Boolean bv;
        @ExcelColumn
        private Character cv;
        @ExcelColumn
        private Short sv;
        @ExcelColumn
        private Integer nv;
        @ExcelColumn
        private Long lv;
        @ExcelColumn
        private Float fv;
        @ExcelColumn
        private Double dv;
        @ExcelColumn
        private String s;
        @ExcelColumn
        private BigDecimal mv;
        @ExcelColumn
        private Date av;
        @ExcelColumn
        private Timestamp iv;
        @ExcelColumn
        private Time tv;
        @ExcelColumn
        private LocalDate ldv;
        @ExcelColumn
        private LocalDateTime ldtv;
        @ExcelColumn
        private LocalTime ltv;

        public static List<BoxAllType> randomTestData(int size) {
            List<BoxAllType> list = new ArrayList<>(size);
            for (int i = 0; i < size; i++) {
                BoxAllType o = new BoxAllType();
                o.bv = random.nextInt(10) == 5;
                o.cv = charArray[random.nextInt(charArray.length)];
                o.sv = (short) (random.nextInt() & 0xFFFF);
                o.nv = random.nextInt();
                o.lv = random.nextLong();
                o.fv = random.nextFloat();
                o.dv = random.nextDouble();
                o.s = getRandomString();
                o.mv = BigDecimal.valueOf(random.nextDouble());
                o.av = new Date();
                o.iv = new Timestamp(System.currentTimeMillis() - random.nextInt(9999999));
                o.tv = new Time(random.nextLong());
                o.ldv = LocalDate.now();
                o.ldtv = LocalDateTime.now();
                o.ltv = LocalTime.now();
                list.add(o);
            }
            return list;
        }

        public static List<BoxAllType> randomTestData() {
            int size = random.nextInt(100) + 1;
            return randomTestData(size);
        }

        public Boolean getBv() {
            return bv;
        }

        public Character getCv() {
            return cv;
        }

        public Short getSv() {
            return sv;
        }

        public Integer getNv() {
            return nv;
        }

        public Long getLv() {
            return lv;
        }

        public Float getFv() {
            return fv;
        }

        public Double getDv() {
            return dv;
        }

        public String getS() {
            return s;
        }

        public BigDecimal getMv() {
            return mv;
        }

        public Date getAv() {
            return av;
        }

        public Timestamp getIv() {
            return iv;
        }

        public Time getTv() {
            return tv;
        }

        public LocalDate getLdv() {
            return ldv;
        }

        public LocalDateTime getLdtv() {
            return ldtv;
        }

        public LocalTime getLtv() {
            return ltv;
        }

        public void setBv(Boolean bv) {
            this.bv = bv;
        }

        public void setCv(Character cv) {
            this.cv = cv;
        }

        public void setSv(Short sv) {
            this.sv = sv;
        }

        public void setNv(Integer nv) {
            this.nv = nv;
        }

        public void setLv(Long lv) {
            this.lv = lv;
        }

        public void setFv(Float fv) {
            this.fv = fv;
        }

        public void setDv(Double dv) {
            this.dv = dv;
        }

        public void setS(String s) {
            this.s = s;
        }

        public void setMv(BigDecimal mv) {
            this.mv = mv;
        }

        public void setAv(Date av) {
            this.av = av;
        }

        public void setIv(Timestamp iv) {
            this.iv = iv;
        }

        public void setTv(Time tv) {
            this.tv = tv;
        }

        public void setLdv(LocalDate ldv) {
            this.ldv = ldv;
        }

        public void setLdtv(LocalDateTime ldtv) {
            this.ldtv = ldtv;
        }

        public void setLtv(LocalTime ltv) {
            this.ltv = ltv;
        }

        @Override
        public boolean equals(Object o) {
            if (this == o) return true;
            if (o == null || getClass() != o.getClass()) return false;
            BoxAllType that = (BoxAllType) o;
            return Objects.equals(bv, that.bv) &&
                Objects.equals(cv, that.cv) &&
                Objects.equals(sv, that.sv) &&
                Objects.equals(nv, that.nv) &&
                Objects.equals(lv, that.lv) &&
                Objects.equals(fv, that.fv) &&
                Objects.equals(dv, that.dv) &&
                Objects.equals(s, that.s) &&
                Objects.equals(mv.setScale(4, RoundingMode.HALF_DOWN), that.mv.setScale(4, RoundingMode.HALF_DOWN)) &&
                av.getTime() / 1000 == that.av.getTime() / 1000 &&
                iv.getTime() / 1000 == that.iv.getTime() / 1000 &&
                String.valueOf(tv).equals(String.valueOf(that.tv)) &&
                Objects.equals(ldv, that.ldv) &&
                Timestamp.valueOf(ldtv).getTime() / 1000 == Timestamp.valueOf(that.ldtv).getTime() / 1000 &&
                String.valueOf(Time.valueOf(ltv)).equals(String.valueOf(Time.valueOf(that.ltv)));
        }

        @Override
        public int hashCode() {
            return Objects.hash(bv, cv, sv, nv, lv, fv, dv, s, mv, av.getTime() / 1000, iv.getTime() / 1000, String.valueOf(tv), ldv, Timestamp.valueOf(ldtv).getTime() / 1000, String.valueOf(Time.valueOf(ltv)));
        }

        @Override
        public String toString() {
            return "" + bv + '|' + cv + '|' + sv + '|' + nv + '|' + lv
                + '|' + fv + '|' + dv + '|' + s + '|' + mv + '|' + av
                + '|' + tv + '|' + ldv + '|' + ldtv + '|' + ltv;
        }
    }

    public static class ExtItem extends Item {
        @ExcelColumn
        private String nice;

        public ExtItem() { }
        public ExtItem(int id, String name) {
            super(id, name);
        }

//        public String getNice() {
//            return nice;
//        }
//
//        public void setNice(String nice) {
//            this.nice = nice;
//        }

        public static List<Item> randomTestData(int n) {
            List<Item> list = new ArrayList<>(n);
            for (int i = 0; i < n; i++) {
                list.add(new ExtItem(i,  getRandomString()));
            }
            return list;
        }

        @Override
        public boolean equals(Object o) {
            if (this == o) return true;
            if (o == null || getClass() != o.getClass()) return false;
            if (!super.equals(o)) return false;
            ExtItem extItem = (ExtItem) o;
            return Objects.equals(nice, extItem.nice);
        }

        @Override
        public int hashCode() {
            return Objects.hash(super.hashCode(), nice);
        }
    }

    public static class NotSharedObject {
        @ExcelColumn(share = false)
        private String name;

        public NotSharedObject() { }
        public NotSharedObject(String name) {
            this.name = name;
        }

        public String getName() {
            return name;
        }

        public void setName(String name) {
            this.name = name;
        }

        @Override
        public boolean equals(Object o) {
            if (this == o) return true;
            if (o == null || getClass() != o.getClass()) return false;
            NotSharedObject that = (NotSharedObject) o;
            return Objects.equals(name, that.name);
        }

        @Override
        public int hashCode() {
            return Objects.hash(name);
        }
    }

    public static class ExtStudent extends Student {
        public ExtStudent() { }
        protected ExtStudent(int id, String name, int score) {
            super(id, name, score);
        }

        @Override
        @ExcelColumn("ID")
        @IgnoreExport
        public int getId() {
            return super.getId();
        }

        @ExcelColumn("ID")
        @Override
        public void setId(int id) {
            super.setId(id);
        }

        @Override
        @ExcelColumn("SCORE")
        @IgnoreExport
        public int getScore() {
            return super.getScore();
        }

        @ExcelColumn("SCORE")
        @Override
        public void setScore(int score) {
            super.setScore(score);
        }
    }
    
    public static class NoColumnAnnotation {
        private int id;
        private String name;

        public int getId() {
            return id;
        }

        public String getName() {
            return name;
        }

        public NoColumnAnnotation() { }
        public NoColumnAnnotation(int id, String name) {
            this.id = id;
            this.name = name;
        }

        public static List<NoColumnAnnotation> randomTestData(int n) {
            List<NoColumnAnnotation> list = new ArrayList<>(n);
            for (int i = 0; i < n; i++) {
                list.add(new NoColumnAnnotation(i, getRandomString()));
            }
            return list;
        }

        public static List<NoColumnAnnotation> randomTestData() {
            int n = random.nextInt(100) + 1;
            return randomTestData(n);
        }

        @Override
        public boolean equals(Object o) {
            if (this == o) return true;
            if (o == null || getClass() != o.getClass()) return false;
            NoColumnAnnotation that = (NoColumnAnnotation) o;
            return id == that.id &&
                Objects.equals(name, that.name);
        }

        @Override
        public int hashCode() {
            return Objects.hash(id, name);
        }

        @Override
        public String toString() {
            return id + " " + name;
        }
    }

    public static class NoColumnAnnotation2 {
        private int age;
        private String abc;

        public int getAge() {
            return age;
        }

        public String getAbc() {
            return abc;
        }

        public NoColumnAnnotation2() { }
        public NoColumnAnnotation2(int age, String abc) {
            this.age = age;
            this.abc = abc;
        }

        public static List<NoColumnAnnotation2> randomTestData(int n) {
            List<NoColumnAnnotation2> list = new ArrayList<>(n);
            for (int i = 0; i < n; i++) {
                list.add(new NoColumnAnnotation2(i, getRandomString()));
            }
            return list;
        }

        public static List<NoColumnAnnotation2> randomTestData() {
            int n = random.nextInt(100) + 1;
            return randomTestData(n);
        }

        @Override
        public boolean equals(Object o) {
            if (this == o) return true;
            if (o == null || getClass() != o.getClass()) return false;
            NoColumnAnnotation2 that = (NoColumnAnnotation2) o;
            return age == that.age &&
                Objects.equals(abc, that.abc);
        }

        @Override
        public int hashCode() {
            return Objects.hash(age, abc);
        }
    }
}

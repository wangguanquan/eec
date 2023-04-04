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

package org.ttzero.excel.entity;

import org.junit.Test;
import org.ttzero.excel.Print;
import org.ttzero.excel.annotation.ExcelColumn;
import org.ttzero.excel.annotation.IgnoreExport;
import org.ttzero.excel.entity.e7.XMLWorksheetWriter;
import org.ttzero.excel.entity.style.Fill;
import org.ttzero.excel.entity.style.Font;
import org.ttzero.excel.entity.style.Horizontals;
import org.ttzero.excel.entity.style.PatternType;
import org.ttzero.excel.entity.style.Styles;
import org.ttzero.excel.manager.Const;
import org.ttzero.excel.manager.docProps.Core;
import org.ttzero.excel.processor.ConversionProcessor;
import org.ttzero.excel.processor.StyleProcessor;
import org.ttzero.excel.reader.Cell;
import org.ttzero.excel.reader.Dimension;
import org.ttzero.excel.reader.ExcelReader;
import org.ttzero.excel.reader.ExcelReaderTest;
import org.ttzero.excel.util.StringUtil;

import java.awt.Color;
import java.beans.IntrospectionException;
import java.beans.Introspector;
import java.beans.PropertyDescriptor;
import java.io.IOException;
import java.lang.reflect.Field;
import java.math.BigDecimal;
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
import java.util.Optional;
import java.util.function.Supplier;
import java.util.stream.Collectors;

import static org.ttzero.excel.Print.println;
import static org.ttzero.excel.reader.Cell.BLANK;
import static org.ttzero.excel.reader.Cell.BOOL;
import static org.ttzero.excel.reader.Cell.CHARACTER;
import static org.ttzero.excel.reader.Cell.DATE;
import static org.ttzero.excel.reader.Cell.DATETIME;
import static org.ttzero.excel.reader.Cell.DECIMAL;
import static org.ttzero.excel.reader.Cell.DOUBLE;
import static org.ttzero.excel.reader.Cell.INLINESTR;
import static org.ttzero.excel.reader.Cell.LONG;
import static org.ttzero.excel.reader.Cell.NUMERIC;
import static org.ttzero.excel.reader.Cell.SST;
import static org.ttzero.excel.reader.Cell.TIME;
import static org.ttzero.excel.reader.ExcelReaderTest.testResourceRoot;

/**
 * @author guanquan.wang at 2019-04-28 19:17
 */
public class ListObjectSheetTest extends WorkbookTest {

    @Test
    public void testWrite() throws IOException {
        new Workbook("test object", author)
            .watch(Print::println)
            .addSheet(Item.randomTestData())
            .writeTo(defaultTestPath);
    }

    @Test public void testAllTypeWrite() throws IOException {
        new Workbook("all type object", author)
            .watch(Print::println)
            .addSheet(AllType.randomTestData())
            .writeTo(defaultTestPath);
    }

    @Test public void testAnnotation() throws IOException {
        new Workbook("annotation object", author)
            .watch(Print::println)
            .addSheet(Student.randomTestData())
            .writeTo(defaultTestPath);
    }

    @Test public void testAnnotationAutoSize() throws IOException {
        new Workbook("annotation object auto-size", author)
            .watch(Print::println)
            .setAutoSize(true)
            .addSheet(new ListSheet<>(Student.randomTestData()))
            .writeTo(defaultTestPath);
    }

    @Test public void testStringWaterMark() throws IOException {
        new Workbook("object string water mark", author)
            .watch(Print::println)
            .setWaterMark(WaterMark.of("SECRET"))
            .addSheet(Item.randomTestData())
            .writeTo(defaultTestPath);
    }

    @Test public void testLocalPicWaterMark() throws IOException {
        new Workbook("object local pic water mark", author)
            .watch(Print::println)
            .setWaterMark(WaterMark.of(testResourceRoot().resolve("mark.png")))
            .addSheet(Item.randomTestData())
            .writeTo(defaultTestPath);
    }

    @Test public void testStreamWaterMark() throws IOException {
        new Workbook("object input stream water mark", author)
            .watch(Print::println)
            .setWaterMark(WaterMark.of(getClass().getClassLoader().getResourceAsStream("mark.png")))
            .addSheet(Item.randomTestData())
            .writeTo(defaultTestPath);
    }

    @Test public void testAutoSize() throws IOException {
        new Workbook("all type auto size", author)
            .watch(Print::println)
            .setAutoSize(true)
            .addSheet(AllType.randomTestData())
            .writeTo(defaultTestPath);
    }

    @Test public void testIntConversion() throws IOException {
        new Workbook("test int conversion", author)
            .watch(Print::println)
            .addSheet(Student.randomTestData()
                , new Column("学号", "id")
                , new Column("姓名", "name")
                , new Column("成绩", "score", n -> (int) n < 60 ? "不及格" : n)
            )
            .writeTo(defaultTestPath);
    }

    @Test public void testStyleConversion() throws IOException {
        new Workbook("object style processor", author)
            .watch(Print::println)
            .addSheet(Student.randomTestData()
                , new Column("学号", "id")
                , new Column("姓名", "name")
                , new Column("成绩", "score")
                    .setStyleProcessor((o, style, sst) -> {
                        if ((int) o < 60) {
                            style = Styles.clearFill(style)
                                | sst.addFill(new Fill(PatternType.solid, Color.orange));
                        }
                        return style;
                    })
            )
            .writeTo(defaultTestPath);
    }

    @Test public void testConvertAndStyleConversion() throws IOException {
        new Workbook("object style and style processor", author)
            .watch(Print::println)
            .addSheet(Student.randomTestData()
                , new Column("学号", "id")
                , new Column("姓名", "name")
                , new Column("成绩", "score", n -> (int) n < 60 ? "不及格" : n)
                    .setStyleProcessor((o, style, sst) -> {
                        if ((int)o < 60) {
                            style = Styles.clearFill(style)
                                | sst.addFill(new Fill(PatternType.solid, new Color(246, 209, 139)));
                        }
                        return style;
                    })
            )
            .writeTo(defaultTestPath);
    }

    @Test public void testCustomizeDataSource() throws IOException {
        new Workbook("customize datasource", author)
            .watch(Print::println)
            .addSheet(new CustomizeDataSourceSheet())
            .writeTo(defaultTestPath);
    }

    @Test public void testBoxAllTypeWrite() throws IOException {
        new Workbook("box all type object", author)
            .watch(Print::println)
            .addSheet(BoxAllType.randomTestData())
            .writeTo(defaultTestPath);
    }

    // -----AUTO SIZE

    @Test public void testBoxAllTypeAutoSizeWrite() throws IOException {
        new Workbook("auto-size box all type object", author)
            .watch(Print::println)
            .setAutoSize(true)
            .addSheet(BoxAllType.randomTestData())
            .writeTo(defaultTestPath);
    }

    @Test public void testCustomizeDataSourceAutoSize() throws IOException {
        new Workbook("auto-size customize datasource", author)
            .watch(Print::println)
            .setAutoSize(true)
            .addSheet(new CustomizeDataSourceSheet())
            .writeTo(defaultTestPath);
    }

    @Test public void testConstructor1() throws IOException {
        new Workbook("test list sheet Constructor1", author)
            .watch(Print::println)
            .setAutoSize(true)
            .addSheet(new ListSheet<Item>())
            .writeTo(defaultTestPath);
    }

    @Test public void testConstructor2() throws IOException {
        new Workbook("test list sheet Constructor2", author)
            .watch(Print::println)
            .setAutoSize(true)
            .addSheet(new ListSheet<Item>("Item").setData(Item.randomTestData(10)))
            .writeTo(defaultTestPath);
    }

    @Test public void testConstructor3() throws IOException {
        new Workbook("test list sheet Constructor3", author)
            .watch(Print::println)
            .setAutoSize(true)
            .addSheet(new ListSheet<Item>("Item"
                , new Column("ID", "id")
                , new Column("NAME", "name"))
                .setData(Item.randomTestData(10)))
            .writeTo(defaultTestPath);
    }

    @Test public void testConstructor4() throws IOException {
        new Workbook("test list sheet Constructor4", author)
            .watch(Print::println)
            .setAutoSize(true)
            .addSheet(new ListSheet<Item>("Item", WaterMark.of(author)
                , new Column("ID", "id")
                , new Column("NAME", "name"))
                .setData(Item.randomTestData(10)))
            .writeTo(defaultTestPath);
    }

    @Test public void testConstructor5() throws IOException {
        new Workbook("test list sheet Constructor5", author)
            .watch(Print::println)
            .setAutoSize(true)
            .addSheet(new ListSheet<>(Item.randomTestData(10)))
            .writeTo(defaultTestPath);
    }

    @Test public void testConstructor6() throws IOException {
        new Workbook("test list sheet Constructor6", author)
            .watch(Print::println)
            .setAutoSize(true)
            .addSheet(new ListSheet<>("ITEM", Item.randomTestData(10)))
            .writeTo(defaultTestPath);
    }

    @Test public void testConstructor7() throws IOException {
        new Workbook("test list sheet Constructor7", author)
            .watch(Print::println)
            .setAutoSize(true)
            .addSheet(new ListSheet<>(Item.randomTestData(10)
                , new Column("ID", "id")
                , new Column("NAME", "name")))
            .writeTo(defaultTestPath);
    }

    @Test public void testConstructor8() throws IOException {
        new Workbook("test list sheet Constructor8", author)
            .watch(Print::println)
            .setAutoSize(true)
            .addSheet(new ListSheet<>("ITEM", Item.randomTestData(10)
                , new Column("ID", "id")
                , new Column("NAME", "name")))
            .writeTo(defaultTestPath);
    }

    @Test public void testConstructor9() throws IOException {
        new Workbook("test list sheet Constructor9", author)
            .watch(Print::println)
            .setAutoSize(true)
            .addSheet(new ListSheet<>(Item.randomTestData(10)
                , WaterMark.of(author)
                , new Column("ID", "id")
                , new Column("NAME", "name")))
            .writeTo(defaultTestPath);
    }

    @Test public void testConstructor10() throws IOException {
        new Workbook("test list sheet Constructor10", author)
            .watch(Print::println)
            .setAutoSize(true)
            .addSheet(new ListSheet<>("ITEM"
                , Item.randomTestData(10)
                , WaterMark.of(author)
                , new Column("ID", "id")
                , new Column("NAME", "name")))
            .writeTo(defaultTestPath);
    }

    @Test public void testArray() throws IOException {
        new Workbook("ListSheet Array as List")
            .watch(Print::println)
            .addSheet(new ListSheet<>()
                .setData(Arrays.asList(new Item(1, "abc"), new Item(2, "xyz"))))
            .writeTo(defaultTestPath);
    }

    @Test public void testSingleList() throws IOException {
        new Workbook("ListSheet Single List")
            .watch(Print::println)
            .addSheet(new ListSheet<>()
                .setData(Collections.singletonList(new Item(1, "a b c"))))
            .writeTo(defaultTestPath);
    }

    public static StyleProcessor sp = (o, style, sst) -> {
        if ((int)o < 60) {
            style = Styles.clearFill(style)
                | sst.addFill(new Fill(PatternType.solid,Color.green, Color.blue));
        }
        return style;
    };

    // 定义一个int值转换lambda表达式，成绩低于60分显示"不及格"，其余显示正常分数
    public static ConversionProcessor conversion = n -> (int) n < 60 ? "不合格" : n;

    @Test
    public void testStyleConversion1() throws IOException {
        new Workbook("2021小五班期未考试成绩")
            .addSheet(new ListSheet<>("期末成绩", Student.randomTestData()
                    , new Column("学号", "id", int.class)
                    , new Column("姓名", "name", String.class)
                    , new Column("成绩", "score", int.class, n -> (int) n < 60 ? "不合格" : n)
                ).setStyleProcessor((o, style, sst) ->
                    o.getScore() < 60 ? Styles.clearFill(style) | sst.addFill(new Fill(PatternType.solid, Color.orange)) : style)
            ).writeTo(defaultTestPath);
    }

    @Test public void testNullValue() throws IOException {
        new Workbook("test null value", author)
            .watch(Print::println)
            .setAutoSize(true)
            .addSheet(new ListSheet<>("EXT-ITEM", ExtItem.randomTestData(10)))
            .writeTo(defaultTestPath);
    }

    @Test public void testReflect() throws IntrospectionException, IllegalAccessException {
        PropertyDescriptor[] array = Introspector.getBeanInfo(ExtItem.class).getPropertyDescriptors();
        for (PropertyDescriptor pd : array) {
            println(pd);
        }
        ExtItem item = new ExtItem(1, "guanquan.wang");
        item.nice = "colvin";

        Field[] fields = item.getClass().getDeclaredFields();
        for (Field field : fields) {
            field.setAccessible(true);
            println(field + ": " + field.get(item));
        }
    }

    @Test public void testFieldUnDeclare() throws IOException {
        new Workbook("field un-declare", author)
            .addSheet(new ListSheet<>("期末成绩", Student.randomTestData()
                    , new Column("学号", "id")
                    , new Column("姓名", "name")
                    , new Column("成绩", "score") // un-declare field
                )
            )
            .writeTo(defaultTestPath);
    }

    @Test public void testResetMethod() throws IOException {
        new Workbook("重写期末成绩", author)
            .addSheet(new ListSheet<Student>("重写期末成绩", Collections.singletonList(new Student(9527, author, 0) {
                    @Override
                    public int getScore() {
                        return 100;
                    }
                }))
            )
            .writeTo(defaultTestPath);
    }

    @Test public void testMethodAnnotation() throws IOException {
        new Workbook("重写方法注解", author)
            .addSheet(new ListSheet<Student>("重写方法注解", Collections.singletonList(new ExtStudent(9527, author, 0) {
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
            .writeTo(defaultTestPath);

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("重写方法注解.xlsx"))) {
            Optional<ExtStudent> opt = reader.sheets().flatMap(org.ttzero.excel.reader.Sheet::dataRows)
                .map(row -> row.too(ExtStudent.class)).findAny();
            assert opt.isPresent();
            ExtStudent student = opt.get();
            assert student.getId() == 9527;
            assert student.getScore() == 0; // The setter column name is 'score'
        }
    }

    @Test public void testIgnoreSupperMethod() throws IOException {
        new Workbook("忽略父类属性", author)
            .addSheet(new ListSheet<Student>("重写方法注解", Collections.singletonList(new ExtStudent(9527, author, 0))))
            .writeTo(defaultTestPath);

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("忽略父类属性.xlsx"))) {
            Optional<ExtStudent> opt = reader.sheets().flatMap(org.ttzero.excel.reader.Sheet::dataRows)
                .map(row -> row.too(ExtStudent.class)).findAny();
            assert opt.isPresent();
            ExtStudent student = opt.get();
            assert student.getId() == 0;
            assert student.getScore() == 0;
        }
    }

    // Issue #93
    @Test public void testListSheet_93() throws IOException {
        new Workbook("Issue#93 List Object").addSheet(new ListSheet<Student>() {
            private int i;
            @Override
            protected List<Student> more() {
                return i++ < 10 ? Student.randomTestData(100) : null;
            }
        }).writeTo(defaultTestPath);
    }

    // Issue #95
    @Test public void testIssue_95() throws IOException {
        new Workbook("Issue #95").addSheet(new ListSheet<NotSharedObject>() {
            private boolean c = true;
            @Override
            protected List<NotSharedObject> more() {
                if (!c) return null;
                c = false;
                List<NotSharedObject> list = new ArrayList<>();
                for (int i = 0; i < 10; i++) {
                    list.add(new NotSharedObject(getRandomString()));
                }
                return list;
            }
        }).writeTo(defaultTestPath);
    }

    @Test public void testSpecifyCore() throws IOException {
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
        new Workbook("Specify Core")
            .setCore(core)
            .addSheet(new ListSheet<>(Collections.singletonList(new NotSharedObject(getRandomString()))))
                .writeTo(defaultTestPath);
    }

    @Test public void testLarge() throws IOException {
        new Workbook("large07").forceExport().addSheet(new ListSheet<ExcelReaderTest.LargeData>() {
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
        }).writeTo(defaultTestPath);

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("large07.xlsx"))) {
            assert Dimension.of("A1:Y50001").equals(reader.sheet(0).getDimension());
        }
    }

    // #132
    @Test public void testEmptyList() throws IOException {
        new Workbook("ListObject empty list").addSheet(new ListSheet<>(new ArrayList<>())).writeTo(defaultTestPath);
    }
    
    @Test public void testNoForceExport() throws IOException {
        new Workbook("testNoForceExport").addSheet(new ListSheet<>(NoColumnAnnotation.randomTestData())).writeTo(defaultTestPath);

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("testNoForceExport.xlsx"))) {
            assert Dimension.of("A1").equals(reader.sheet(0).getDimension());
        }
    }
    
    @Test public void testForceExportOnWorkbook() throws IOException {
        int lines = random.nextInt(100) + 3;
        new Workbook("testForceExportOnWorkbook").forceExport().addSheet(new ListSheet<>(NoColumnAnnotation.randomTestData(lines))).writeTo(defaultTestPath);
        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("testForceExportOnWorkbook.xlsx"))) {
            assert reader.sheet(0).dataRows().count() == lines;
        }
    }

    @Test public void testForceExportOnWorkSheet() throws IOException {
        int lines = random.nextInt(100) + 3;
        new Workbook("testForceExportOnWorkSheet").addSheet(new ListSheet<>(NoColumnAnnotation.randomTestData(lines)).forceExport()).writeTo(defaultTestPath);
        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("testForceExportOnWorkSheet.xlsx"))) {
            assert reader.sheet(0).dataRows().count() == lines;
        }
    }

    @Test public void testForceExportOnWorkbook2() throws IOException {
        int lines = random.nextInt(100) + 3, lines2 = random.nextInt(100) + 4;
        new Workbook("testForceExportOnWorkbook2")
                .forceExport()
                .addSheet(new ListSheet<>(NoColumnAnnotation.randomTestData(lines)))
                .addSheet(new ListSheet<>(NoColumnAnnotation2.randomTestData(lines2)))
                .writeTo(defaultTestPath);
        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("testForceExportOnWorkbook2.xlsx"))) {
            assert reader.sheet(0).dataRows().count() == lines;
            assert reader.sheet(1).dataRows().count() == lines2;
        }
    }

    @Test public void testForceExportOnWorkbook2Cancel1() throws IOException {
        int lines = random.nextInt(100) + 3, lines2 = random.nextInt(100) + 4;
        new Workbook("testForceExportOnWorkbook2Cancel1")
                .forceExport()
                .addSheet(new ListSheet<>(NoColumnAnnotation.randomTestData(lines)).cancelForceExport())
                .addSheet(new ListSheet<>(NoColumnAnnotation2.randomTestData(lines2)))
                .writeTo(defaultTestPath);
        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("testForceExportOnWorkbook2Cancel1.xlsx"))) {
            assert reader.sheets().count() == 2L;
            assert reader.sheet(0).dataRows().count() == 0L;
            assert reader.sheet(1).dataRows().count() == lines2;
        }
    }

    @Test public void testForceExportOnWorkbook2Cancel2() throws IOException {
        int lines = random.nextInt(100) + 3, lines2 = random.nextInt(100) + 4;
        new Workbook("testForceExportOnWorkbook2Cancel2")
                .forceExport()
                .addSheet(new ListSheet<>(NoColumnAnnotation.randomTestData(lines)).cancelForceExport())
                .addSheet(new ListSheet<>(NoColumnAnnotation2.randomTestData(lines2)).cancelForceExport())
                .writeTo(defaultTestPath);
        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("testForceExportOnWorkbook2Cancel2.xlsx"))) {
            assert reader.sheets().count() == 2L;
            assert reader.sheet(0).dataRows().count() == 0L;
            assert reader.sheet(1).dataRows().count() == 0L;
        }
    }

    @Test public void testWrapText() throws IOException {
        new Workbook("WRAP TEXT")
                .addSheet(new ListSheet<>()
                        .setData(Arrays.asList(new Item(1, "a b c\r\n1 2 3\r\n中文\t测试\r\nAAAAAA")
                                , new Item(2, "fdsafdsafdsafdsafdsafdsafdsafdsfadsafdsafdsafdsafdsafdsafds"))))
                .writeTo(defaultTestPath);
    }

    @Test public void testClearHeadStyle() throws IOException {
        Workbook workbook = new Workbook("clear style").addSheet(new ListSheet<>(Item.randomTestData()));

        Sheet sheet = workbook.getSheetAt(0);
        sheet.cancelOddStyle();  // Clear odd style
        int headStyle = sheet.defaultHeadStyle();
        sheet.setHeadStyle(Styles.clearFill(headStyle) & Styles.clearFont(headStyle));
        sheet.setHeadStyle(sheet.getHeadStyle() | workbook.getStyles().addFont(new Font("宋体", 11, Font.Style.BOLD, Color.BLACK)));

        workbook.writeTo(defaultTestPath);
    }



    @Test public void testBasicType() throws IOException {
        List<Integer> list = new ArrayList<>(35);
        for (int i = 0; i < 35; i++) list.add(i);
        new Workbook()
            .addSheet(new ListSheet<>(list).setColumns(new ListSheet.EntryColumn().setClazz(Integer.class).setCellStyle(0)))
            .writeTo(defaultTestPath.resolve("Integer array.xlsx"));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("Integer array.xlsx"))) {
            Integer[] array = reader.sheets().flatMap(org.ttzero.excel.reader.Sheet::rows).map(row -> row.getInt(0)).toArray(Integer[]::new);
            assert array.length == list.size();
            for (int i = 0; i < array.length; i++) {
                assert array[i].equals(list.get(i));
            }
        }
    }

    @Test public void testEmojiChar() throws IOException {
        List<String> list = Arrays.asList("😂", "abc😍(●'◡'●)cz");
        new Workbook()
            .addSheet(new ListSheet<>(list).setColumns(new ListSheet.EntryColumn().setClazz(String.class)))
            .writeTo(defaultTestPath.resolve("Emoji char.xlsx"));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("Emoji char.xlsx"))) {
            List<String> subList = reader.sheet(0).rows().map(row -> row.getString(0)).collect(Collectors.toList());

            assert subList.size() == list.size();

            for (int i = 0, len = subList.size(); i < len; i++) {
                assert subList.get(i).equals(list.get(i));
            }
        }
    }

    @Test public void test264() throws IOException {
        Column[] columns = {new Column("ID", "id"), new Column("NAME", "name")};
        new Workbook().addSheet(new ListSheet<>(Item.randomTestData(10), columns))
            .writeTo(defaultTestPath.resolve("Issue 264.xlsx"));

        new Workbook().addSheet(new ListSheet<>(Item.randomTestData(10), columns))
            .writeTo(defaultTestPath.resolve("Issue 264-1.xlsx"));

        new Workbook().addSheet(new ListSheet<>(Item.randomTestData(10), columns))
            .writeTo(defaultTestPath.resolve("Issue 264-2.xlsx"));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("Issue 264-1.xlsx"))) {
            Iterator<org.ttzero.excel.reader.Row> iter = reader.sheet(0).iterator();
            if (iter.hasNext()) {
                org.ttzero.excel.reader.Row row = iter.next();
                assert "ID".equals(row.getString(0));
                assert "NAME".equals(row.getString(1));
            }
        }

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("Issue 264-2.xlsx"))) {
            Iterator<org.ttzero.excel.reader.Row> iter = reader.sheet(0).iterator();
            if (iter.hasNext()) {
                org.ttzero.excel.reader.Row row = iter.next();
                assert "ID".equals(row.getString(0));
                assert "NAME".equals(row.getString(1));
            }
        }
    }

    @Test public void testNullInList() throws IOException {
        List<Item> list = Item.randomTestData(10);
        list.add(0, null);
        list.add(3, null);
        list.add(null);
        new Workbook("Null in list").addSheet(new ListSheet<>(list)).writeTo(defaultTestPath);
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
                Student e = new Student(i, (k = random.nextInt(10)) < 3 ? new String(new char[] {(char)('a' + k)}) : getRandomString(), random.nextInt(50) + 50);
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
        @ExcelColumn
        public String toString() {
            return "id: " + id + ", name: " + name + ", score: " + score;
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

        public static List<AllType> randomTestData(int size) {
            List<AllType> list = new ArrayList<>(size);
            for (int i = 0; i < size; i++) {
                AllType o = new AllType();
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
        public String toString() {
            return "" + bv + '|' + cv + '|' + sv + '|' + nv + '|' + lv
                + '|' + fv + '|' + dv + '|' + s + '|' + mv + '|' + av
                + '|' + tv + '|' + ldv + '|' + ldtv + '|' + ltv;
        }
    }

    public static class ExtItem extends Item {
        @ExcelColumn
        private String nice;

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
    }

    public static class NotSharedObject {
        @ExcelColumn(share = false)
        private String name;

        public NotSharedObject(String name) {
            this.name = name;
        }

        public String getName() {
            return name;
        }

        public void setName(String name) {
            this.name = name;
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
    }

    public static class Template {
        @ExcelColumn(maxWidth = 12.0D, wrapText = true)
        String v1;
        @ExcelColumn(maxWidth = 20.0, wrapText = true)
        String v2;
        @ExcelColumn(maxWidth = 25.0D, wrapText = true)
        String v3;

        static Template of(String v1, String v2, String v3) {
            Template v = new Template();
            v.v1 = v1;
            v.v2 = v2;
            v.v3 = v3;
            return v;
        }
    }

    @Test public void testCustomerRowHeight() throws IOException {
        List<Template> list = new ArrayList<>();
        list.add(Template.of("备注说明\r\n第二行\r\n第三行\r\n第四行", "岗位名称", "岁位"));
        list.add(Template.of("字段名称", "*岗位名称", "岗位描述"));
        list.add(Template.of("示例", "生产统计员", "按照产品规格、价格、工序、员工、车间等不同对象和要求进行统计数据资料分析"));

        new Workbook("Customer row height").addSheet(
            new ListSheet<>(list).setStyleProcessor(new TemplateStyleProcessor())
                .cancelOddStyle().ignoreHeader().putExtProp(Const.ExtendPropertyKey.MERGE_CELLS, Collections.singletonList(Dimension.of("A1:B1")))
                .setSheetWriter(new XMLWorksheetWriter() {
                    protected int startRow(int rows, int columns, double rowHeight) throws IOException {
                        // Row number
                        int r = rows + startRow;
                        // logging
                        if (r % 1_0000 == 0) {
                            sheet.what("0014", String.valueOf(r));
                        }

                        bw.write("<row r=\"");
                        bw.writeInt(r);
                        // default data row height 16.5
                        bw.write("\" spans=\"1:");
                        bw.writeInt(columns);
                        bw.write("\" ht=\"62.25\" customHeight=\"1\">");
                        return r;
                    }
                })
        ).writeTo(defaultTestPath);
    }

    public static class TemplateStyleProcessor implements StyleProcessor<Template> {
        String k;
        int c = 0;
        @Override
        public int build(Template o, int style, Styles sst) {
            if (!o.v1.equals(k)) {
                k = o.v1;
                c = 0;
            }
            switch (o.v1) {
                case "备注说明":
                    if (c > 0)
                        style = Styles.clearFill(style) | sst.addFill(new Fill(PatternType.solid, new Color(188, 219, 162)));
                    break;
                case "字段名称":
                    Font font = sst.getFont(style);
                    style = Styles.clearFont(style) | sst.addFont(font.clone().bold());
                    if (c > 0)
                        style = Styles.clearHorizontal(style) | Horizontals.CENTER;
                    break;
                case "示例":
                    if (c == 1)
                        style = Styles.clearHorizontal(style) | Horizontals.CENTER;
                    break;
            }
            c++;
            return style;
        }
    }


    @Test public void testTileWriter() throws IOException {
        List<TileEntity> data = TileEntity.randomTestData();
        new Workbook("Dynamic title").cancelOddFill().addSheet(new ListSheet<>(data).setSheetWriter(new TileXMLWorksheetWriter(3, LocalDate.now().toString()))).writeTo(defaultTestPath);
    }

    public static class TileEntity {
        @ExcelColumn("{date} 拣货单")
        @ExcelColumn(value = "差异", maxWidth = 8.6D)
        private String diff;
        @ExcelColumn("{date} 拣货单")
        @ExcelColumn(value = "序号", maxWidth = 6.8D)
        private Integer no;
        @ExcelColumn("{date} 拣货单")
        @ExcelColumn(value = "商品", maxWidth = 12.0D)
        private String product;
        @ExcelColumn("{date} 拣货单")
        @ExcelColumn(value = "数量", maxWidth = 6.8D)
        private Integer num;

        public static List<TileEntity> randomTestData() {
            int n = 23;
            List<TileEntity> list = new ArrayList<>(n);
            for (int i = 0; i < n; i++) {
                TileEntity e = new TileEntity();
                e.no = i + 1;
                e.product = getRandomString(10);
                e.num = random.nextInt(20) + 1;
                list.add(e);
            }
            return list;
        }
    }

    /**
     * 自定义平铺WorksheetWriter
     */
    public static class TileXMLWorksheetWriter extends XMLWorksheetWriter {
        private int tile; // 平铺的数量，也就是每行重复输出多少条数据
        private String date; // 可忽略，仅仅是表头上的日期

        public TileXMLWorksheetWriter(int tile) {
            this.tile = tile;
        }

        public TileXMLWorksheetWriter(int tile, String date) {
            this.tile = tile;
            this.date = date;
        }

        public int getTile() {
            return tile;
        }

        public void setTile(int tile) {
            this.tile = tile;
        }

        public String getDate() {
            return date;
        }

        public void setDate(String date) {
            this.date = date;
        }

        @Override
        protected void writeBefore() throws IOException {
            // The header columns
            columns = sheet.getAndSortHeaderColumns();
            // Give new columns
            tileColumns();

            boolean nonHeader = sheet.getNonHeader() == 1;

            bw.write(Const.EXCEL_XML_DECLARATION);
            // Declaration
            bw.newLine();
            // Root node
            writeRootNode();

            // Dimension
            writeDimension();

            // SheetViews default value
            writeSheetViews();

            // Default row height and width
            int fillSpace = 6;
            BigDecimal width = BigDecimal.valueOf(!nonHeader ? sheet.getDefaultWidth() : 8.38D);
            String defaultWidth = width.setScale(2, BigDecimal.ROUND_HALF_UP).toString();
            writeSheetFormat(fillSpace, defaultWidth);

            // cols
            writeCols(fillSpace, defaultWidth);

            // Write body data
            beforeSheetData(nonHeader);
        }

        protected void tileColumns() {
            if (tile == 1) return;

            int x = columns.length, y = x * tile, t = columns[columns.length - 1].getRealColIndex();
            // Bound check
            if (y > Const.Limit.MAX_COLUMNS_ON_SHEET)
                throw new TooManyColumnsException(y, Const.Limit.MAX_COLUMNS_ON_SHEET);

            Column[] _columns = new Column[y];
            for (int i = 0; i < y; i++) {
                // 第一个对象的表头不需要复制
                Column col = i < x ? columns[i] : new Column(columns[i % x]).addSubColumn(new Column());
                col.realColIndex = columns[i % x].realColIndex + t * (i / x);
                _columns[i] = col;

                // 替换拣货单上的日期
                Column _col = col;
                do {
                    if (StringUtil.isNotEmpty(_col.getName()) && _col.getName().contains("{date}"))
                        _col.setName(_col.getName().replace("{date}", date));
                }
                while ((_col = _col.next) != null);
            }

            this.columns = _columns;

            // FIXME 这里强行指定合并替换掉原本的头
            List<Dimension> mergeCells = Collections.singletonList(new Dimension(1, (short) 1, 1, (short) y));
            sheet.putExtProp(Const.ExtendPropertyKey.MERGE_CELLS, mergeCells);
        }

        @Override
        protected void writeRow(Row row) throws IOException {
            Cell[] cells = row.getCells();
            int len = cells.length, r = row.getIndex() / tile + startRow, c = columns[columns.length - 1].realColIndex / tile, y = row.getIndex() % tile;
            if (y == 0) startRow(r - startRow, columns[columns.length - 1].realColIndex, -1D);

            for (int i = 0; i < len; i++) {
                Cell cell = cells[i];
                int xf = cell.xf, col = i + c * y;
                switch (cell.t) {
                    case INLINESTR:
                    case SST:
                        writeString(cell.sv, r, col, xf);
                        break;
                    case NUMERIC:
                        writeNumeric(cell.nv, r, col, xf);
                        break;
                    case LONG:
                        writeNumeric(cell.lv, r, col, xf);
                        break;
                    case DATE:
                    case DATETIME:
                    case DOUBLE:
                    case TIME:
                        writeDouble(cell.dv, r, col, xf);
                        break;
                    case BOOL:
                        writeBool(cell.bv, r, col, xf);
                        break;
                    case DECIMAL:
                        writeDecimal(cell.mv, r, col, xf);
                        break;
                    case CHARACTER:
                        writeChar(cell.cv, r, col, xf);
                        break;
                    case BLANK:
                        writeNull(r, col, xf);
                        break;
                    default:
                }
            }
            // 注意这里可能不会关闭row需要在writeAfter进行二次处理
            if (y == tile - 1)
                bw.write("</row>");
        }

        @Override
        protected void writeAfter(int total) throws IOException {
            if (total > 0 && (total - 1) % tile < tile - 1) bw.write("</row>");
            super.writeAfter(total);
        }
    }
}

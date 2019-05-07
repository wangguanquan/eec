/*
 * Copyright (c) 2019, guanquan.wang@yandex.com All Rights Reserved.
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

package cn.ttzero.excel.entity;

import cn.ttzero.excel.Print;
import cn.ttzero.excel.annotation.DisplayName;
import cn.ttzero.excel.annotation.NotExport;
import cn.ttzero.excel.entity.style.Fill;
import cn.ttzero.excel.entity.style.PatternType;
import cn.ttzero.excel.entity.style.Styles;
import cn.ttzero.excel.reader.ExcelReaderTest;
import org.junit.Test;

import java.awt.*;
import java.io.IOException;
import java.math.BigDecimal;
import java.sql.Time;
import java.sql.Timestamp;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

/**
 * Create by guanquan.wang at 2019-04-28 19:17
 */
public class ListObjectSheetTest extends WorkbookTest{

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
            .setWaterMark(WaterMark.of(ExcelReaderTest.testResourceRoot().resolve("mark.png")))
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
                , new Sheet.Column("学号", "id", int.class)
                , new Sheet.Column("姓名", "name", String.class)
                , new Sheet.Column("年龄", "age", int.class, n -> n > 14 ? "高龄" : n)
            )
            .writeTo(defaultTestPath);
    }

    @Test public void testStyleConversion() throws IOException {
        new Workbook("object style processor", author)
            .watch(Print::println)
            .addSheet(Student.randomTestData()
                , new Sheet.Column("学号", "id", int.class)
                , new Sheet.Column("姓名", "name", String.class)
                , new Sheet.Column("年龄", "age", int.class)
                    .setStyleProcessor((o, style, sst) -> {
                        int n = (int) o;
                        if (n < 10) {
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
                , new Sheet.Column("学号", "id", int.class)
                , new Sheet.Column("姓名", "name", String.class)
                , new Sheet.Column("年龄", "age", int.class, n -> n > 14 ? "高龄" : n)
                    .setStyleProcessor((o, style, sst) -> {
                        int n = (int) o;
                        if (n > 14) {
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
                , new Sheet.Column("ID", "id", int.class)
                , new Sheet.Column("NAME", "name", String.class))
                .setData(Item.randomTestData(10)))
            .writeTo(defaultTestPath);
    }

    @Test public void testConstructor4() throws IOException {
        new Workbook("test list sheet Constructor4", author)
            .watch(Print::println)
            .setAutoSize(true)
            .addSheet(new ListSheet<Item>("Item", WaterMark.of(author)
                , new Sheet.Column("ID", "id", int.class)
                , new Sheet.Column("NAME", "name", String.class))
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
                , new Sheet.Column("ID", "id", int.class)
                , new Sheet.Column("NAME", "name", String.class)))
            .writeTo(defaultTestPath);
    }

    @Test public void testConstructor8() throws IOException {
        new Workbook("test list sheet Constructor8", author)
            .watch(Print::println)
            .setAutoSize(true)
            .addSheet(new ListSheet<>("ITEM", Item.randomTestData(10)
                , new Sheet.Column("ID", "id", int.class)
                , new Sheet.Column("NAME", "name", String.class)))
            .writeTo(defaultTestPath);
    }

    @Test public void testConstructor9() throws IOException {
        new Workbook("test list sheet Constructor9", author)
            .watch(Print::println)
            .setAutoSize(true)
            .addSheet(new ListSheet<>(Item.randomTestData(10)
                , WaterMark.of(author)
                , new Sheet.Column("ID", "id", int.class)
                , new Sheet.Column("NAME", "name", String.class)))
            .writeTo(defaultTestPath);
    }

    @Test public void testConstructor10() throws IOException {
        new Workbook("test list sheet Constructor10", author)
            .watch(Print::println)
            .setAutoSize(true)
            .addSheet(new ListSheet<>("ITEM"
                , Item.randomTestData(10)
                , WaterMark.of(author)
                , new Sheet.Column("ID", "id", int.class)
                , new Sheet.Column("NAME", "name", String.class)))
            .writeTo(defaultTestPath);
    }

    public static class Item {
        private int id;
        private String name;

        Item(int id, String name) {
            this.id = id;
            this.name = name;
        }

        public static List<Item> randomTestData(int n) {
            List<Item> list = new ArrayList<>(n);
            for (int i = 0; i < n; i++) {
                list.add(new Item(i, getRandomString()));
            }
            return list;
        }

        public static List<Item> randomTestData() {
            int n = random.nextInt(100) + 1;
            return randomTestData(n);
        }
    }

    public static class AllType {
        private boolean bv;
        private char cv;
        private short sv;
        private int nv;
        private long lv;
        private float fv;
        private double dv;
        private String s;
        private BigDecimal mv;
        private Date av;
        private Timestamp iv;
        private Time tv;
        private LocalDate ldv;
        private LocalDateTime ldtv;
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
    }

    /**
     * Annotation Object
     */
    public static class Student {
        @DisplayName("学号")
        private int id;
        @DisplayName("姓名")
        private String name;
        @DisplayName("年龄")
        private int age;
        @NotExport("secret")
        private String password;

        Student(int id, String name, int age) {
            this.id = id;
            this.name = name;
            this.age = age;
        }
        public static List<Student> randomTestData(int pageNo, int limit) {
            List<Student> list = new ArrayList<>(limit);
            for (int i = pageNo * limit, n = i + limit; i < n; i++) {
                Student e = new Student(i, getRandomString(), random.nextInt(15) + 5);
                e.password = getRandomString();
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
    }

    public static class BoxAllType {
        private Boolean bv;
        private Character cv;
        private Short sv;
        private Integer nv;
        private Long lv;
        private Float fv;
        private Double dv;
        private String s;
        private BigDecimal mv;
        private Date av;
        private Timestamp iv;
        private Time tv;
        private LocalDate ldv;
        private LocalDateTime ldtv;
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
    }
}

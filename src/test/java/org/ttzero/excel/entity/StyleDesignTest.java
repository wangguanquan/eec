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
import org.ttzero.excel.annotation.ExcelColumn;
import org.ttzero.excel.annotation.StyleDesign;
import org.ttzero.excel.entity.style.Border;
import org.ttzero.excel.entity.style.BorderStyle;
import org.ttzero.excel.entity.style.Fill;
import org.ttzero.excel.entity.style.Font;
import org.ttzero.excel.entity.style.Horizontals;
import org.ttzero.excel.entity.style.PatternType;
import org.ttzero.excel.entity.style.Styles;
import org.ttzero.excel.manager.Const;
import org.ttzero.excel.processor.StyleProcessor;
import org.ttzero.excel.reader.Dimension;
import org.ttzero.excel.reader.ExcelReader;

import java.awt.Color;
import java.io.IOException;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.stream.Collectors;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;

/**
 * @author guanquan.wang at 2022-07-15 23:31
 */
public class StyleDesignTest extends WorkbookTest {

    @Test public void testStyleDesign() throws IOException {
        new Workbook()
            .addSheet(new ListSheet<>("期末成绩", DesignStudent.randomTestData()))
            .writeTo(defaultTestPath.resolve("标识行样式.xlsx"));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("标识行样式.xlsx"))) {
            reader.sheet(0).header(1).bind(DesignStudent.class).rows().forEach(row -> {
                Styles styles = row.getStyles();
                DesignStudent o = row.get();
                int c0 = row.getCellStyle(0), c1 = row.getCellStyle(1), c2 = row.getCellStyle(2);
                Fill f0 = styles.getFill(c0), f1 = styles.getFill(c1), f2 = styles.getFill(c2);
                if (o.getScore() < 60) {
                    assertTrue(f0 != null && f0.getPatternType() == PatternType.solid && f0.getFgColor().equals(Color.orange));
                    assertTrue(f1 != null && f1.getPatternType() == PatternType.solid && f1.getFgColor().equals(Color.orange));
                    assertTrue(f2 != null && f2.getPatternType() == PatternType.solid && f2.getFgColor().equals(Color.orange));
                } else if (o.getScore() < 70) {
                    assertTrue(f0 != null && f0.getPatternType() == PatternType.solid && f0.getFgColor().equals(Color.green));
                    assertTrue(f1 != null && f1.getPatternType() == PatternType.solid && f1.getFgColor().equals(Color.green));
                    assertTrue(f2 != null && f2.getPatternType() == PatternType.solid && f2.getFgColor().equals(Color.green));
                } else if (o.getScore() > 90) {
                    Font ft0 = styles.getFont(c0), ft1 = styles.getFont(c1), ft2 = styles.getFont(c2);
                    assertTrue(ft0.isUnderline() && ft0.isBold());
                    assertTrue(ft1.isUnderline() && ft0.isBold());
                    assertTrue(ft2.isUnderline() && ft0.isBold());
                } else {
                    assertTrue(f0 == null || f0.getPatternType() == PatternType.none);
                    assertTrue(f1 == null || f1.getPatternType() == PatternType.none);
                    assertTrue(f2 == null || f2.getPatternType() == PatternType.none);
                }

                if (VIP_SET.contains(o.getName())) {
                    Font font = styles.getFont(c0);
                    assertTrue(font.isBold());
                }
            });
        }
    }

    @Test public void testStyleDesign1() throws IOException {
        ListSheet<ListObjectSheetTest.Item> itemListSheet = new ListSheet<>("序列数", ListObjectSheetTest.Item.randomTestData());
        itemListSheet.setStyleProcessor(rainbowStyle);
        new Workbook().addSheet(itemListSheet).writeTo(defaultTestPath.resolve("标识行样式1.xlsx"));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("标识行样式1.xlsx"))) {
            reader.sheet(0).header(1).bind(ListObjectSheetTest.Item.class).rows().forEach(row -> {
                Styles styles = row.getStyles();
                ListObjectSheetTest.Item item = row.get();
                int c0 = row.getCellStyle(0), c1 = row.getCellStyle(1);
                Fill f0 = styles.getFill(c0), f1 = styles.getFill(c1);
                if (item.getId() % 3 == 0) {
                    assertTrue(f0 != null && f0.getPatternType() == PatternType.solid && f0.getFgColor().equals(Color.green));
                    assertTrue(f1 != null && f1.getPatternType() == PatternType.solid && f1.getFgColor().equals(Color.green));
                } else if (item.getId() % 3 == 1) {
                    assertTrue(f0 != null && f0.getPatternType() == PatternType.solid && f0.getFgColor().equals(Color.blue));
                    assertTrue(f1 != null && f1.getPatternType() == PatternType.solid && f1.getFgColor().equals(Color.blue));
                } else if (item.getId() % 3 == 2) {
                    assertTrue(f0 != null && f0.getPatternType() == PatternType.solid && f0.getFgColor().equals(Color.pink));
                    assertTrue(f1 != null && f1.getPatternType() == PatternType.solid && f1.getFgColor().equals(Color.pink));
                } else {
                    assertTrue(f0 == null || f0.getPatternType() == PatternType.none);
                    assertTrue(f1 == null || f1.getPatternType() == PatternType.none);
                }
            });
        }
    }

    @Test public void testStyleDesign2() throws IOException {
        new Workbook()
            .addSheet(new ListSheet<>("序列数", DesignStudent.randomTestData()).setStyleProcessor((item, style, sst) -> {
                if (item != null && item.getId() < 10) {
                    style = sst.modifyFill(style, new Fill(PatternType.solid, Color.green));
                }
                return style;
            }))
            .writeTo(defaultTestPath.resolve("标识行样式2.xlsx"));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("标识行样式2.xlsx"))) {
            reader.sheet(0).header(1).bind(DesignStudent.class).rows().forEach(row -> {
                Styles styles = row.getStyles();
                DesignStudent item = row.get();
                int c0 = row.getCellStyle(0), c1 = row.getCellStyle(1), c2 = row.getCellStyle(2);
                Fill f0 = styles.getFill(c0), f1 = styles.getFill(c1), f2 = styles.getFill(c2);
                if (item.getId() < 10) {
                    assertTrue(f0 != null && f0.getPatternType() == PatternType.solid && f0.getFgColor().equals(Color.green));
                    assertTrue(f1 != null && f1.getPatternType() == PatternType.solid && f1.getFgColor().equals(Color.green));
                    assertTrue(f2 != null && f2.getPatternType() == PatternType.solid && f2.getFgColor().equals(Color.green));
                } else {
                    assertTrue(f0 == null || f0.getPatternType() == PatternType.none);
                    assertTrue(f1 == null || f1.getPatternType() == PatternType.none);
                    assertTrue(f2 == null || f2.getPatternType() == PatternType.none);
                }

                if (VIP_SET.contains(item.getName())) {
                    Font font = styles.getFont(c0);
                    assertTrue(font.isBold());
                }
            });
        }
    }

    @Test public void testStyleDesignSpecifyColumns() throws IOException {
        new Workbook()
            .addSheet(new ListSheet<>("序列数", DesignStudent.randomTestData()
                , new Column("姓名", "name").setWrapText(true).setStyleProcessor((n, s, sst) -> sst.modifyHorizontal(s, Horizontals.CENTER))
                , new Column("数学成绩", "score").setWidth(12D)
                , new Column("备注", "toString").setWidth(25.32D).setWrapText(true)
            )).writeTo(defaultTestPath.resolve("标识行样式3.xlsx"));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("标识行样式3.xlsx"))) {
            reader.sheet(0).header(1).rows().forEach(row -> {
                Styles styles = row.getStyles();
                int c0 = row.getCellStyle(0), c2 = row.getCellStyle(2);
                assertEquals(styles.getWrapText(c0), 1);
                assertEquals(styles.getHorizontal(c0), Horizontals.CENTER);
                assertEquals(styles.getWrapText(c2), 1);
            });
        }
    }

    @Test public void testMergedCells() throws IOException {
        List<E> list = new ArrayList<>();
        list.add(new E("暗月月", "男", "1", "数学", "3", 30, "教育a", "教育b"));
        list.add(new E("暗月月", "男", "2", "语文", "1", 30, "教育a", "教育c"));
        list.add(new E("暗月月", "男", "3", "历史", "1", 30, "教育b", "教育c"));
        list.add(new E("张三", "女", "1", "英语", "1", 20, "教育d", "教育d"));
        list.add(new E("张三", "女", "5", "物理", "7", 20, "教育x", "教育x"));
        list.add(new E("李四", "男", "2", "语文", "1", 24, "教育c", "教育a"));
        list.add(new E("李四", "男", "3", "历史", "1", 24, "教育b", "教育c"));
        list.add(new E("王五", "男", "1", "高数", "2", 28, "教育c", "教育a"));
        list.add(new E("王五", "男", "2", "JAvA", "3", 28, "教育b", "教育c"));

        List<Dimension> mergeCells = new ArrayList<>();
        String name = null;
        int row = 2, nameFrom = row;
        for (E e : list) {
            if (!e.name.equals(name)) {
                if (row > nameFrom + 1) {
                    mergeCells.add(new Dimension(nameFrom + 1, (short) 1, row, (short) 1));
                    mergeCells.add(new Dimension(nameFrom + 1, (short) 2, row, (short) 2));
                    mergeCells.add(new Dimension(nameFrom + 1, (short) 6, row, (short) 6));
                }
                name = e.name;
                nameFrom = row;
            } else {
                e.name = null;
                e.sex = null;
                e.age = null;
            }
            row++;
        }
        if (row > nameFrom + 1) {
            mergeCells.add(new Dimension(nameFrom + 1, (short) 1, row, (short) 1));
            mergeCells.add(new Dimension(nameFrom + 1, (short) 2, row, (short) 2));
            mergeCells.add(new Dimension(nameFrom + 1, (short) 6, row, (short) 6));
        }
        new Workbook().cancelZebraLine().addSheet(new LightListSheet<>(list
            , new Column("姓名", "name")
            , new Column("性别", "sex")
            , new Column("证书").addSubColumn(new Column("编号", "no"))
            , new Column("证书").addSubColumn(new Column("类型", "type"))
            , new Column("证书").addSubColumn(new Column("等级", "level"))
            , new Column("年龄", "age")
            , new Column("教育").addSubColumn(new Column("教育1", "jy1"))
            , new Column("教育").addSubColumn(new Column("教育2", "jy2")))
            .setStyleProcessor(new GroupStyleProcessor<>())
            .putExtProp(Const.ExtendPropertyKey.MERGE_CELLS, mergeCells))
            .writeTo(defaultTestPath.resolve("Merged Cells.xlsx"));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("Merged Cells.xlsx"))) {
            // Normal reader
            List<Map<String, Object>> readList = reader.sheet(0).header(1, 2).rows().map(org.ttzero.excel.reader.Row::toMap).collect(Collectors.toList());
            List<String> expected = Arrays.asList("{姓名=暗月月, 性别=男, 证书:编号=1, 证书:类型=数学, 证书:等级=3, 年龄=30, 教育:教育1=教育a, 教育:教育2=教育b}",
                "{姓名=, 性别=, 证书:编号=2, 证书:类型=语文, 证书:等级=1, 年龄=, 教育:教育1=教育a, 教育:教育2=教育c}",
                "{姓名=, 性别=, 证书:编号=3, 证书:类型=历史, 证书:等级=1, 年龄=, 教育:教育1=教育b, 教育:教育2=教育c}",
                "{姓名=张三, 性别=女, 证书:编号=1, 证书:类型=英语, 证书:等级=1, 年龄=20, 教育:教育1=教育d, 教育:教育2=教育d}",
                "{姓名=, 性别=, 证书:编号=5, 证书:类型=物理, 证书:等级=7, 年龄=, 教育:教育1=教育x, 教育:教育2=教育x}",
                "{姓名=李四, 性别=男, 证书:编号=2, 证书:类型=语文, 证书:等级=1, 年龄=24, 教育:教育1=教育c, 教育:教育2=教育a}",
                "{姓名=, 性别=, 证书:编号=3, 证书:类型=历史, 证书:等级=1, 年龄=, 教育:教育1=教育b, 教育:教育2=教育c}",
                "{姓名=王五, 性别=男, 证书:编号=1, 证书:类型=高数, 证书:等级=2, 年龄=28, 教育:教育1=教育c, 教育:教育2=教育a}",
                "{姓名=, 性别=, 证书:编号=2, 证书:类型=JAvA, 证书:等级=3, 年龄=, 教育:教育1=教育b, 教育:教育2=教育c}");

            assertEquals(readList.size(), expected.size());
            for (int i = 0; i < expected.size(); i++) {
                assertEquals(expected.get(i), readList.get(i).toString());
            }

            // Copy on merged reader
            List<Map<String, Object>> readList2 = reader.sheet(0).asFullSheet().copyOnMerged().header(1, 2).rows().map(org.ttzero.excel.reader.Row::toMap).collect(Collectors.toList());
            List<String> expected2 = Arrays.asList("{姓名=暗月月, 性别=男, 证书:编号=1, 证书:类型=数学, 证书:等级=3, 年龄=30, 教育:教育1=教育a, 教育:教育2=教育b}",
                "{姓名=暗月月, 性别=男, 证书:编号=2, 证书:类型=语文, 证书:等级=1, 年龄=30, 教育:教育1=教育a, 教育:教育2=教育c}",
                "{姓名=暗月月, 性别=男, 证书:编号=3, 证书:类型=历史, 证书:等级=1, 年龄=30, 教育:教育1=教育b, 教育:教育2=教育c}",
                "{姓名=张三, 性别=女, 证书:编号=1, 证书:类型=英语, 证书:等级=1, 年龄=20, 教育:教育1=教育d, 教育:教育2=教育d}",
                "{姓名=张三, 性别=女, 证书:编号=5, 证书:类型=物理, 证书:等级=7, 年龄=20, 教育:教育1=教育x, 教育:教育2=教育x}",
                "{姓名=李四, 性别=男, 证书:编号=2, 证书:类型=语文, 证书:等级=1, 年龄=24, 教育:教育1=教育c, 教育:教育2=教育a}",
                "{姓名=李四, 性别=男, 证书:编号=3, 证书:类型=历史, 证书:等级=1, 年龄=24, 教育:教育1=教育b, 教育:教育2=教育c}",
                "{姓名=王五, 性别=男, 证书:编号=1, 证书:类型=高数, 证书:等级=2, 年龄=28, 教育:教育1=教育c, 教育:教育2=教育a}",
                "{姓名=王五, 性别=男, 证书:编号=2, 证书:类型=JAvA, 证书:等级=3, 年龄=28, 教育:教育1=教育b, 教育:教育2=教育c}");

            assertEquals(readList2.size(), expected2.size());
            for (int i = 0; i < expected2.size(); i++) {
                assertEquals(expected2.get(i), readList2.get(i).toString());
            }
        }
    }

    public static class E implements Group {
        private String name, sex, no, type, level, jy1, jy2;
        private Integer age;

        public E(String name, String sex, String no, String type, String level, Integer age, String jy1, String jy2) {
            this.name = name;
            this.sex = sex;
            this.no = no;
            this.type = type;
            this.level = level;
            this.age = age;
            this.jy1 = jy1;
            this.jy2 = jy2;
        }

        @Override
        public String groupBy() {
            return name;
        }
    }

    // =======================公共部分=======================
    public interface Group {
        String groupBy();
    }

    public static class GroupStyleProcessor<U extends Group> implements StyleProcessor<U> {
        private String group;
        private int s, o;
        @Override
        public int build(U u, int style, Styles sst) {
            if (group == null) {
                group = u.groupBy();
                s = sst.addFill(new Fill(PatternType.solid, new Color(233, 234, 236)));
                return style;
            }
            if (u.groupBy() != null && !group.equals(u.groupBy())) {
                group = u.groupBy();
                o ^= 1;
            }
            return o == 1 ? Styles.clearFill(style) | s : style;
        }
    }

    public static class LightListSheet<T> extends ListSheet<T> {
        public LightListSheet(List<T> data, Column... columns) {
            super(data, columns);
        }

        @Override
        protected int init() {
            int v = super.init();
            Styles styles = workbook.getStyles();
            try {
                Field borderField = Styles.class.getDeclaredField("borders");
                borderField.setAccessible(true);
                @SuppressWarnings("unchecked")
                List<Border> borders = (List<Border>) borderField.get(styles);
                if (borders != null && borders.size() > 1) {
                    Border border = borders.get(1);
                    border.setBorder(BorderStyle.THIN, new Color(191, 191, 191));
                }

                Field fontField = Styles.class.getDeclaredField("fonts");
                fontField.setAccessible(true);
                @SuppressWarnings("unchecked")
                List<Font> fonts = (List<Font>) fontField.get(styles);
                if (fonts != null && fonts.size() > 1) {
                    fonts.get(0).setName("Microsoft JhengHei");
                    fonts.get(1).setName("Microsoft JhengHei");
                }
            } catch (NoSuchFieldException | IllegalAccessException e) {
                // Ignore
            }
            return v;
        }
    }

    public static class StudentScoreStyle implements StyleProcessor<DesignStudent> {
        @Override
        public int build(DesignStudent o, int style, Styles sst) {
            // 低于60分时背景色标黄
            if (o.getScore() < 60) {
                style = sst.modifyFill(style, new Fill(PatternType.solid, Color.orange));
            } else if (o.getScore() < 70) {
                style = sst.modifyFill(style, new Fill(PatternType.solid, Color.green));
            } else if (o.getScore() > 90) {
                Font newFont = sst.getFont(style).clone();
                style = sst.modifyFont(style, newFont.underline().bold());
            }
            return style;
        }
    }

    public static StyleProcessor<ListObjectSheetTest.Item> rainbowStyle = (item, style, sst) -> {
        if (item.getId() % 3 == 0) {
            style = sst.modifyFill(style, new Fill(PatternType.solid, Color.green));
        } else if (item.getId() % 3 == 1) {
            style = sst.modifyFill(style, new Fill(PatternType.solid, Color.blue));
        } else if (item.getId() % 3 == 2) {
            style = sst.modifyFill(style, new Fill(PatternType.solid, Color.pink));
        }
        return style;
    };


    private static final Set<String> VIP_SET = new HashSet<>(Arrays.asList("a", "b", "x"));

    public static class NameMatch implements StyleProcessor<String> {
        @Override
        public int build(String name, int style, Styles sst) {
            if (VIP_SET.contains(name)) {
                Font font = sst.getFont(style).clone();
                style = sst.modifyFont(style, font.bold());
            }
            return style;
        }
    }

    /**
     * Annotation Object
     */
    @StyleDesign(using = StudentScoreStyle.class)
    public static class DesignStudent extends ListObjectSheetTest.Student {

        @ExcelColumn("姓名")
        @StyleDesign(using = NameMatch.class)
        @Override
        public String getName() {
            return super.getName();
        }

        @ExcelColumn
        public int getCid() {
            return super.getId();
        }

        @ExcelColumn
        public void setCid(int id) {
            super.setId(id);
        }

        public DesignStudent() { }

        public DesignStudent(int id, String name, int score) {
            super(id, name, score);
        }

        public static List<ListObjectSheetTest.Student> randomTestData(int pageNo, int limit) {
            List<ListObjectSheetTest.Student> list = new ArrayList<>(limit);
            for (int i = pageNo * limit, n = i + limit, k; i < n; i++) {
                ListObjectSheetTest.Student e = new DesignStudent(i, (k = random.nextInt(10)) < 3 ? new String(new char[]{(char) ('a' + k)}) : getRandomString(), random.nextInt(50) + 50);
                list.add(e);
            }
            return list;
        }

        public static List<ListObjectSheetTest.Student> randomTestData(int n) {
            return randomTestData(0, n);
        }

        public static List<ListObjectSheetTest.Student> randomTestData() {
            int n = random.nextInt(100) + 1;
            return randomTestData(n);
        }
    }
}

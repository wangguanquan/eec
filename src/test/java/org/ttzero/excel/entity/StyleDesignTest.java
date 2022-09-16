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

import java.awt.Color;
import java.io.IOException;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashSet;
import java.util.List;
import java.util.Set;

/**
 * @author guanquan.wang at 2022-07-15 23:31
 */
public class StyleDesignTest extends WorkbookTest {

    @Test
    public void testStyleDesign() throws IOException {
        new Workbook("标识行样式", author)
            .addSheet(new ListSheet<>("期末成绩", DesignStudent.randomTestData()))
            .writeTo(defaultTestPath);
    }

    @Test
    public void testStyleDesign1() throws IOException {
        ListSheet<ListObjectSheetTest.Item> itemListSheet = new ListSheet<>("序列数", ListObjectSheetTest.Item.randomTestData());
        itemListSheet.setStyleProcessor(rainbowStyle);
        new Workbook("标识行样式1", author)
            .addSheet(itemListSheet)
            .writeTo(defaultTestPath);
    }

    @Test
    public void testStyleDesign2() throws IOException {
        new Workbook("标识行样式2", author)
            .addSheet(new ListSheet<>("序列数", DesignStudent.randomTestData()).setStyleProcessor((item, style, sst, axis) -> {
                if (item != null && item.getId() < 10) {
                    style = Styles.clearFill(style) | sst.addFill(new Fill(PatternType.solid, Color.green));
                }
                return style;
            }))
            .writeTo(defaultTestPath);
    }

    @Test
    public void testStyleDesignSpecifyColumns() throws IOException {
        new Workbook("标识行样式3", author)
            .addSheet(new ListSheet<>("序列数", DesignStudent.randomTestData()
                , new Column("姓名", "name").setWrapText(true).setStyleProcessor((n, s, sst, axis) -> Styles.clearHorizontal(s) | Horizontals.CENTER)
                , new Column("数学成绩", "score").setWidth(12D)
                , new Column("备注", "toString").setWidth(25.32D).setWrapText(true)
            )).writeTo(defaultTestPath);
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
        new Workbook("Merged Cells").cancelOddFill().addSheet(new LightListSheet<>(list
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
            .writeTo(defaultTestPath);
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
        public int build(U u, int style, Styles sst, Axis axis) {
            if (group == null) {
                group = u.groupBy();
                s = sst.addFill(new Fill(PatternType.solid, new Color(239, 245, 235)));
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
        public LightListSheet(List<T> data, org.ttzero.excel.entity.Column... columns) {
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
        public int build(DesignStudent o, int style, Styles sst, Axis axis) {
            // 低于60分时背景色标黄
            if (o.getScore() < 60) {
                style = Styles.clearFill(style) | sst.addFill(new Fill(PatternType.solid, Color.orange));
                // 低于30分时加下划线
            } else if (o.getScore() < 70) {
                style = Styles.clearFill(style) | sst.addFill(new Fill(PatternType.solid, Color.green));
            } else if (o.getScore() > 90) {
                // 获取原有字体+下划线（这样做可以保留原字体和大小）
                Font newFont = sst.getFont(style).clone();
                style = Styles.clearFont(style) | sst.addFont(newFont.underLine().bold());
            }
            return style;
        }
    }

    public static StyleProcessor<ListObjectSheetTest.Item> rainbowStyle = (item, style, sst, axis) -> {
        if (item.getId() % 3 == 0) {
            style = Styles.clearFill(style) | sst.addFill(new Fill(PatternType.solid, Color.green));
        } else if (item.getId() % 3 == 1) {
            style = Styles.clearFill(style) | sst.addFill(new Fill(PatternType.solid, Color.blue));
        } else if (item.getId() % 3 == 2) {
            style = Styles.clearFill(style) | sst.addFill(new Fill(PatternType.solid, Color.pink));
        }
        return style;
    };


    private static final Set<String> VIP_SET = new HashSet<>(Arrays.asList("a", "b", "x"));

    public static class NameMatch implements StyleProcessor<String> {
        @Override
        public int build(String name, int style, Styles sst, Axis axis) {
            if (VIP_SET.contains(name)) {
                Font font = sst.getFont(style).clone();
                style = Styles.clearFont(style) | sst.addFont(font.bold());
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

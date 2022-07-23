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
import org.ttzero.excel.entity.style.Fill;
import org.ttzero.excel.entity.style.Font;
import org.ttzero.excel.entity.style.Horizontals;
import org.ttzero.excel.entity.style.PatternType;
import org.ttzero.excel.entity.style.Styles;
import org.ttzero.excel.processor.StyleProcessor;

import java.awt.Color;
import java.io.IOException;
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
            .addSheet(new ListSheet<>("序列数", DesignStudent.randomTestData()).setStyleProcessor((item, style, sst) -> {
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
                , new Column("姓名", "name").setWrapText(true).setStyleProcessor((n, s, sst) -> Styles.clearHorizontal(s) | Horizontals.CENTER)
                , new Column("数学成绩", "score").setWidth(12D)
                , new Column("备注", "toString").setWidth(25.32D).setWrapText(true)
            )).writeTo(defaultTestPath);
    }


    public static class StudentScoreStyle implements StyleProcessor<DesignStudent> {
        @Override
        public int build(DesignStudent o, int style, Styles sst) {
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

    public static StyleProcessor<ListObjectSheetTest.Item> rainbowStyle = (item, style, sst) -> {
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
        public int build(String name, int style, Styles sst) {
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

/*
 * Copyright (c) 2017-2020, guanquan.wang@yandex.com All Rights Reserved.
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
import org.ttzero.excel.annotation.HeaderComment;
import org.ttzero.excel.annotation.ExcelColumn;
import org.ttzero.excel.entity.style.Font;
import org.ttzero.excel.reader.ExcelReader;

import java.awt.Color;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.stream.Collectors;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;


/**
 * @author guanquan.wang at 2020-05-21 16:52
 */
public class CommentTest extends WorkbookTest {
    @Test public void testComment() throws IOException {
        String fileName = "comment test.xlsx";
        List<Student> expectList = Student.randomTestData();
        new Workbook()
            .addSheet(new ListSheet<>(expectList))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            List<Student> list = reader.sheet(0).dataRows().map(row -> row.to(Student.class)).collect(Collectors.toList());
            assertEquals(expectList.size(), list.size());
            for (int i = 0, len = expectList.size(); i < len; i++) {
                Student expect = expectList.get(i), e = list.get(i);
                assertEquals(expect, e);
            }

            Map<Long, Comment> commentMap = reader.sheet(0).asFullSheet().getComments();
            Comment A1 = commentMap.get(1L << 16 | 1);
            assertNotNull(A1);
            assertEquals(A1.title, "王老师：");
            assertEquals(A1.value, "学生ID");

            Comment B1 = commentMap.get(1L << 16 | 2);
            assertNotNull(B1);
            assertEquals(B1.title, "王老师：");
            assertEquals(B1.value, "学生姓名");

            Comment C1 = commentMap.get(1L << 16 | 3);
            assertNotNull(C1);
            assertEquals(C1.value, "低于60分显示\"不合格\"");
        }
    }

    @Test public void testCommentLongText() throws IOException {
        String fileName = "long text comment test.xlsx";
        List<Student> expectList = Student.randomTestData();
        Sheet sheet = new ListSheet<>(expectList);
        Comments comments = sheet.getComments();
        if (comments == null) comments = sheet.createComments();
        comments.addComment("C5", new Comment("提示：", "1、第一行批注内容较多时无法完全显示内容，增加弹出框大小设置\n" +
            "2、第二行批注内容较多时无法完全显示内容\n" +
            "3、第三行批注内容较多时无法完全显示内容\n" +
            "4、第四行批注内容较多时无法完全显示内容", 180D, 80D));
        new Workbook()
            .addSheet(sheet)
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            List<Student> list = reader.sheet(0).dataRows().map(row -> row.to(Student.class)).collect(Collectors.toList());
            assertEquals(expectList.size(), list.size());
            for (int i = 0, len = expectList.size(); i < len; i++) {
                Student expect = expectList.get(i), e = list.get(i);
                assertEquals(expect, e);
            }

            Map<Long, Comment> commentMap = reader.sheet(0).asFullSheet().getComments();
            Comment C5 = commentMap.get(5L << 16 | 3);
            assertNotNull(C5);
            assertEquals(C5.title, "提示：");
            assertEquals(C5.value, "1、第一行批注内容较多时无法完全显示内容，增加弹出框大小设置\n" +
                "2、第二行批注内容较多时无法完全显示内容\n" +
                "3、第三行批注内容较多时无法完全显示内容\n" +
                "4、第四行批注内容较多时无法完全显示内容");
        }
    }

    @Test public void testBodyComments() throws IOException {
        final String fileName = "Body批注测试.xlsx";
        List<Stock> list = new ArrayList<>();
        list.add(new Stock(60, StockHealth.HEALTHY));
        list.add(new Stock(40, StockHealth.NORMAL));
        list.add(new Stock(10, StockHealth.DANGER));

        Font yh10 = new Font("微软雅黑", 10, Color.RED);
        Workbook workbook = new Workbook();
        ListSheet<Stock> listSheet = new ListSheet<>(list);
        // 获取Comments
        Comments comments = listSheet.createComments();
        // A1单元格添加批注
        comments.addComment(1, 1, "实际库存");
        // B4单元格添加批注
        comments.addComment(4, 2, new Comment("库存不足", "低于警戒线13%,请尽快添加库存").setValueFont(yh10));
        workbook.addSheet(listSheet);
        workbook.writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            Map<Long, Comment> commentMap = reader.sheet(0).asFullSheet().getComments();
            Comment A1 = commentMap.get(1L << 16 | 1);
            assertNotNull(A1);
            assertEquals(A1.value, "实际库存");

            Comment B4 = commentMap.get(4L << 16 | 2);
            assertNotNull(B4);
            assertEquals(B4.title, "库存不足");
            assertEquals(B4.value, "低于警戒线13%,请尽快添加库存");
            assertEquals(yh10, B4.valueFont);
        }
    }

    /**
     * Annotation Object
     */
    public static class Student {
        @ExcelColumn(value = "SID", comment = @HeaderComment(title = "王老师：", value = "学生ID"))
        private int id;
        @ExcelColumn(value = "姓名", comment = @HeaderComment(title = "王老师：", value = "学生姓名"))
        private String name;
        @HeaderComment(title = "王老师：", value = "低于60分显示\"不合格\"")
        @ExcelColumn(value = "成绩")
        private int score;

        public Student() { }

        public Student(int id, String name, int score) {
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
            for (int i = pageNo * limit, n = i + limit; i < n; i++) {
               Student e = new Student(i, getRandomString(), random.nextInt(50) + 50);
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
        public boolean equals(Object o) {
            if (this == o) return true;
            if (o == null || getClass() != o.getClass()) return false;
            Student student = (Student) o;
            return id == student.id &&
                score == student.score &&
                Objects.equals(name, student.name);
        }

        @Override
        public int hashCode() {
            return Objects.hash(id, name, score);
        }

        @Override
        public String toString() {
            return "id: " + id + ", name: " + name + ", score: " + score;
        }
    }

    public static class Stock {
        @ExcelColumn("库存")
        private int stock;
        @ExcelColumn(value = "库存健康度", comment = @HeaderComment(value =
            "健康：库存大于阈值20%\n" +
            "正常：库存高于阈值10%\n" +
            "警告：库存高于阈值0～10%\n" +
            "危险：库存低于阈值10%", width = 120))
        private StockHealth stockHealth;

        public Stock(int stock, StockHealth stockHealth) {
            this.stock = stock;
            this.stockHealth = stockHealth;
        }
    }

    public enum StockHealth {
        HEALTHY("健康", 2),
        NORMAL("正常", 1),
        WARNING("警告", 0),
        DANGER("危险", -1);

        StockHealth(String desc, int threshold) {
            this.desc = desc;
            this.threshold = threshold;
        }

        private final String desc;
        private final int threshold;

        public String getDesc() {
            return desc;
        }

        @Override
        public String toString() {
            return desc;
        }
    }
}

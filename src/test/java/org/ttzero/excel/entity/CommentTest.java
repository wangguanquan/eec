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
import org.ttzero.excel.reader.ExcelReader;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.stream.Collectors;

import static org.junit.Assert.assertEquals;


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
}

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
import org.ttzero.excel.manager.Const;

import java.io.IOException;
import java.sql.Timestamp;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

/**
 * @author guanquan.wang at 2022-07-24 10:34
 */
public class CustomColIndexTest extends WorkbookTest {

    @Test
    public void testOrderColumn() throws IOException {
        new Workbook(("Order column")).addSheet(new ListSheet<>(OrderEntry.randomTestData())).writeTo(defaultTestPath);
    }

    @Test public void testSameOrderColumn() throws IOException {
        new Workbook(("Same order column")).addSheet(new ListSheet<>(SameOrderEntry.randomTestData())).writeTo(defaultTestPath);
    }

    @Test public void testFractureOrderColumn() throws IOException {
        new Workbook(("Fracture order column")).addSheet(new ListSheet<>(FractureOrderEntry.randomTestData())).writeTo(defaultTestPath);
    }

    @Test public void testLargeOrderColumn() throws IOException {
        new Workbook(("Large order column")).addSheet(new ListSheet<>(LargeOrderEntry.randomTestData())).writeTo(defaultTestPath);
    }

    @Test public void testOverLargeOrderColumn() throws IOException {
        try {
            new Workbook(("Over Large order column")).addSheet(new ListSheet<>(OverLargeOrderEntry.randomTestData())).writeTo(defaultTestPath);
            assert false;
        } catch (TooManyColumnsException e) {
            assert true;
        }
    }

    @Test public void testOrderColumnSpecifyOnColumn() throws IOException {
        new Workbook("Order column 2")
            .addSheet(new ListSheet<>("期末成绩", ListObjectSheetTest.Student.randomTestData()
                , new Column("学号", "id").setColIndex(3)
                , new Column("姓名", "name")
                , new Column("成绩", "score").setColIndex(5) // un-declare field
            )).writeTo(defaultTestPath);
    }


    public static class OrderEntry {
        @ExcelColumn(colIndex = 0)
        private String s;
        @ExcelColumn( colIndex = 1)
        private Date date;
        @ExcelColumn(colIndex = 2)
        private Double d;
        @ExcelColumn(colIndex = 3)
        private String s2 = "a";
        @ExcelColumn(colIndex = 4)
        private String s3 = "b";
        @ExcelColumn(colIndex = 5)
        private String s4 = "c";

        public OrderEntry() {}
        public OrderEntry(String s, Date date, Double d) {
            this.s = s;
            this.date = date;
            this.d = d;
        }

        public static List<OrderEntry> randomTestData(int n) {
            List<OrderEntry> list = new ArrayList<>(n);
            for (int i = 0; i < n; i++) {
                list.add(new OrderEntry(getRandomString(), new Timestamp(System.currentTimeMillis() - random.nextInt(9999999)), (double) i));
            }
            return list;
        }

        public static List<OrderEntry> randomTestData() {
            int n = random.nextInt(100) + 1;
            return randomTestData(n);
        }

        public String getS() {
            return s;
        }

        public Date getDate() {
            return date;
        }

        public Double getD() {
            return d;
        }

        public String getS2() {
            return s2;
        }

        public String getS3() {
            return s3;
        }

        public String getS4() {
            return s4;
        }
    }

    public static class SameOrderEntry extends OrderEntry {
        public SameOrderEntry() {}
        public SameOrderEntry(String s, Date date, Double d) {
            super(s, date, d);
        }

        @Override
        @ExcelColumn(colIndex = 5)
        public Double getD() {
            return super.getD();
        }

        @Override
        @ExcelColumn(colIndex = 5)
        public String getS2() {
            return super.getS2();
        }

        public static List<OrderEntry> randomTestData(int n) {
            List<OrderEntry> list = new ArrayList<>(n);
            for (int i = 0; i < n; i++) {
                list.add(new SameOrderEntry(getRandomString(), new Timestamp(System.currentTimeMillis() - random.nextInt(9999999)), (double) i));
            }
            return list;
        }

        public static List<OrderEntry> randomTestData() {
            int n = random.nextInt(100) + 1;
            return randomTestData(n);
        }
    }

    public static class FractureOrderEntry extends OrderEntry {
        public FractureOrderEntry() {}
        public FractureOrderEntry(String s, Date date, Double d) {
            super(s, date, d);
        }

        @Override
        @ExcelColumn
        public String getS() {
            return super.getS();
        }

        @Override
        @ExcelColumn
        public Date getDate() {
            return super.getDate();
        }

        @Override
        @ExcelColumn(colIndex = 2)
        public Double getD() {
            return super.getD();
        }

        @Override
        @ExcelColumn(colIndex = 0)
        public String getS2() {
            return super.getS2();
        }

        @Override
        @ExcelColumn
        public String getS3() {
            return super.getS3();
        }

        @Override
        @ExcelColumn(colIndex = 4)
        public String getS4() {
            return super.getS4();
        }

        public static List<OrderEntry> randomTestData(int n) {
            List<OrderEntry> list = new ArrayList<>(n);
            for (int i = 0; i < n; i++) {
                list.add(new FractureOrderEntry(getRandomString(), new Timestamp(System.currentTimeMillis() - random.nextInt(9999999)), (double) i));
            }
            return list;
        }

        public static List<OrderEntry> randomTestData() {
            int n = random.nextInt(100) + 1;
            return randomTestData(n);
        }
    }

    public static class LargeOrderEntry extends OrderEntry {
        public LargeOrderEntry() {}
        public LargeOrderEntry(String s, Date date, Double d) {
            super(s, date, d);
        }

        @Override
        @ExcelColumn(colIndex = 1)
        public String getS() {
            return super.getS();
        }

        @Override
        @ExcelColumn(colIndex = Const.Limit.MAX_COLUMNS_ON_SHEET - 1)
        public Date getDate() {
            return super.getDate();
        }

        @Override
        @ExcelColumn(colIndex = 189)
        public String getS2() {
            return super.getS2();
        }

        public static List<OrderEntry> randomTestData(int n) {
            List<OrderEntry> list = new ArrayList<>(n);
            for (int i = 0; i < n; i++) {
                list.add(new LargeOrderEntry(getRandomString(), new Timestamp(System.currentTimeMillis() - random.nextInt(9999999)), (double) i));
            }
            return list;
        }

        public static List<OrderEntry> randomTestData() {
            int n = random.nextInt(100) + 1;
            return randomTestData(n);
        }
    }

    public static class OverLargeOrderEntry extends OrderEntry {
        public OverLargeOrderEntry() {}
        public OverLargeOrderEntry(String s, Date date, Double d) {
            super(s, date, d);
        }

        @Override
        @ExcelColumn(colIndex = 16_384)
        public String getS() {
            return super.getS();
        }

        public static List<OrderEntry> randomTestData(int n) {
            List<OrderEntry> list = new ArrayList<>(n);
            for (int i = 0; i < n; i++) {
                list.add(new OverLargeOrderEntry(getRandomString(), new Timestamp(System.currentTimeMillis() - random.nextInt(9999999)), (double) i));
            }
            return list;
        }

        public static List<OrderEntry> randomTestData() {
            int n = random.nextInt(100) + 1;
            return randomTestData(n);
        }
    }
}

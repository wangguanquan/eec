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
import org.ttzero.excel.reader.ExcelReader;
import org.ttzero.excel.reader.HeaderRow;

import java.io.IOException;
import java.sql.Timestamp;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.List;
import java.util.Objects;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;

/**
 * @author guanquan.wang at 2022-07-24 10:34
 */
public class CustomColIndexTest extends WorkbookTest {

    @Test public void testOrderColumn() throws IOException {
        String fileName = "Order column.xlsx";
        List<OrderEntry> expectList = OrderEntry.randomTestData();
        new Workbook().addSheet(new ListSheet<>(expectList)).writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            Iterator<org.ttzero.excel.reader.Row> iter = reader.sheet(0).header(1).iterator();
            for (OrderEntry expect : expectList) {
                assertTrue(iter.hasNext());
                org.ttzero.excel.reader.Row row = iter.next();
                assertEquals(expect, row.to(OrderEntry.class));
            }
        }
    }

    @Test public void testSameOrderColumn() throws IOException {
        String fileName = "Same order column.xlsx";
        List<OrderEntry> expectList = SameOrderEntry.randomTestData();
        new Workbook().addSheet(new ListSheet<>(expectList)).writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            org.ttzero.excel.reader.Sheet sheet = reader.sheet(0).header(1);
            org.ttzero.excel.reader.HeaderRow header = (HeaderRow) sheet.getHeader();
            assertEquals("s", header.get(0));
            assertEquals("date", header.get(1));
            assertEquals("s3", header.get(4));
            assertEquals("d", header.get(5));
            assertEquals("s2", header.get(6));
            assertEquals("s4", header.get(7));

            Iterator<org.ttzero.excel.reader.Row> iter = sheet.iterator();
            for (OrderEntry expect : expectList) {
                assertTrue(iter.hasNext());
                org.ttzero.excel.reader.Row row = iter.next();
                assertEquals(expect.s, row.getString("s"));
                assertEquals(expect.date.getTime() / 1000, row.getTimestamp("date").getTime() / 1000);
                assertEquals(Double.compare(expect.d, row.getDouble("d")), 0);
                assertEquals(expect.s2, row.getString("s2"));
                assertEquals(expect.s4, row.getString("s4"));
            }
        }
    }

    @Test public void testFractureOrderColumn() throws IOException {
        String fileName = "Fracture order column.xlsx";
        List<OrderEntry> expectList = FractureOrderEntry.randomTestData();
        new Workbook().addSheet(new ListSheet<>(expectList)).writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            org.ttzero.excel.reader.Sheet sheet = reader.sheet(0).header(1);
            org.ttzero.excel.reader.HeaderRow header = (HeaderRow) sheet.getHeader();
            assertEquals("s2", header.get(0));
            assertEquals("s", header.get(1));
            assertEquals("d", header.get(2));
            assertEquals("date", header.get(3));
            assertEquals("s4", header.get(4));
            assertEquals("s3", header.get(5));

            Iterator<org.ttzero.excel.reader.Row> iter = sheet.iterator();
            for (OrderEntry expect : expectList) {
                assertTrue(iter.hasNext());
                org.ttzero.excel.reader.Row row = iter.next();
                assertEquals(expect.s2, row.getString("s2"));
                assertEquals(expect.s, row.getString("s"));
                assertEquals(Double.compare(expect.d, row.getDouble("d")), 0);
                assertEquals(expect.date.getTime() / 1000, row.getTimestamp("date").getTime() / 1000);
                assertEquals(expect.s4, row.getString("s4"));
                assertEquals(expect.s3, row.getString("s3"));
            }
        }
    }

    @Test public void testLargeOrderColumn() throws IOException {
        String fileName = "Large order column.xlsx";
        List<OrderEntry> expectList = LargeOrderEntry.randomTestData();
        new Workbook().addSheet(new ListSheet<>(expectList)).writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            org.ttzero.excel.reader.Sheet sheet = reader.sheet(0).header(1);
            org.ttzero.excel.reader.HeaderRow header = (HeaderRow) sheet.getHeader();
            assertEquals("s", header.get(1));
            assertEquals("d", header.get(2));
            assertEquals("s3", header.get(4));
            assertEquals("s4", header.get(5));
            assertEquals("s2", header.get(189));
            assertEquals("date", header.get(Const.Limit.MAX_COLUMNS_ON_SHEET - 1));

            Iterator<org.ttzero.excel.reader.Row> iter = sheet.iterator();
            for (OrderEntry expect : expectList) {
                assertTrue(iter.hasNext());
                org.ttzero.excel.reader.Row row = iter.next();
                assertEquals(expect.s, row.getString("s"));
                assertEquals(Double.compare(expect.d, row.getDouble("d")), 0);
                assertEquals(expect.s3, row.getString("s3"));
                assertEquals(expect.s4, row.getString("s4"));
                assertEquals(expect.s2, row.getString("s2"));
                assertEquals(expect.date.getTime() / 1000, row.getTimestamp("date").getTime() / 1000);
            }
        }
    }

    @Test(expected = TooManyColumnsException.class) public void testOverLargeOrderColumn() throws IOException {
        new Workbook(("Over Large order column")).addSheet(new ListSheet<>(OverLargeOrderEntry.randomTestData())).writeTo(defaultTestPath);
    }

    @Test public void testOrderColumnSpecifyOnColumn() throws IOException {
        String fileName = "Order column 2.xlsx";
        List<ListObjectSheetTest.Student> expectList = ListObjectSheetTest.Student.randomTestData();
        new Workbook()
            .addSheet(new ListSheet<>("期末成绩", expectList
                , new Column("学号", "id").setColIndex(3)
                , new Column("姓名", "name")
                , new Column("成绩", "score").setColIndex(5)
            )).writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            org.ttzero.excel.reader.Sheet sheet = reader.sheet(0).header(1);
            assertEquals("期末成绩", sheet.getName());
            org.ttzero.excel.reader.HeaderRow header = (HeaderRow) sheet.getHeader();
            assertEquals("姓名", header.get(0));
            assertEquals("学号", header.get(3));
            assertEquals("成绩", header.get(5));

            Iterator<org.ttzero.excel.reader.Row> iter = sheet.iterator();
            for (ListObjectSheetTest.Student expect : expectList) {
                assertTrue(iter.hasNext());
                org.ttzero.excel.reader.Row row = iter.next();
                ListObjectSheetTest.Student e = row.too(ListObjectSheetTest.Student.class);
                expect.setId(0); // ID ignore field
                assertEquals(expect, e);
            }
        }
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

        @Override
        public boolean equals(Object o) {
            if (this == o) return true;
            if (o == null || getClass() != o.getClass()) return false;
            OrderEntry that = (OrderEntry) o;
            return Objects.equals(s, that.s) &&
                date.getTime() / 1000 == that.date.getTime() / 1000 &&
                Double.compare(d, that.d) == 0 &&
                Objects.equals(s2, that.s2) &&
                Objects.equals(s3, that.s3) &&
                Objects.equals(s4, that.s4);
        }

        @Override
        public int hashCode() {
            return Objects.hash(s, date, d, s2, s3, s4);
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

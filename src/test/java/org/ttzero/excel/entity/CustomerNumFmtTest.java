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
import org.ttzero.excel.annotation.ExcelColumn;
import org.ttzero.excel.entity.style.NumFmt;
import org.ttzero.excel.entity.style.Styles;
import org.ttzero.excel.reader.ExcelReader;
import org.ttzero.excel.reader.HeaderRow;

import java.io.IOException;
import java.sql.Timestamp;
import java.text.NumberFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.function.Supplier;

import static org.ttzero.excel.util.ExtBufferedWriter.stringSize;

/**
 * @author guanquan.wang at 2020-09-30 10:34
 */
public class CustomerNumFmtTest extends WorkbookTest {

    @Test public void testStringSize() {
        assert 4 == stringSize(1234);
        assert 5 == stringSize(-1234);
        assert 16 == stringSize(1231234354837485L);
        assert 17 == stringSize(-1231234354837485L);
    }

    @Test public void testFmtInAnnotation() throws IOException {
        String fileName = "customize_numfmt.xlsx";
        List<Item> expectList = Item.random();
        new Workbook().setAutoSize(true).addSheet(new ListSheet<>(expectList)).writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            Iterator<org.ttzero.excel.reader.Row> iter = reader.sheet(0).header(1).bind(Item.class).iterator();
            for (Item expect : expectList) {
                assert iter.hasNext();
                org.ttzero.excel.reader.Row row = iter.next();
                assert expect.equals(row.get());

                Styles styles = row.getStyles();
                int styleIndex = row.getCellStyle(2);
                NumFmt numFmt = styles.getNumFmt(styleIndex);
                assert numFmt != null && "yyyy\\-mm\\-dd".equals(numFmt.getCode());
            }
        }
    }

    @Test public void testFmtInAnnotationYmdHms() throws IOException {
        String fileName = "customize_numfmt_full.xlsx";
        List<ItemFull> expectList = ItemFull.randomFull();
        new Workbook().setAutoSize(true).addSheet(new ListSheet<>(expectList)).writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            Iterator<org.ttzero.excel.reader.Row> iter = reader.sheet(0).header(1).bind(ItemFull.class).iterator();
            for (ItemFull expect : expectList) {
                assert iter.hasNext();
                org.ttzero.excel.reader.Row row = iter.next();
                assert expect.equals(row.get());

                Styles styles = row.getStyles();
                int styleIndex = row.getCellStyle(3);
                NumFmt numFmt = styles.getNumFmt(styleIndex);
                assert numFmt != null && "yyyy\\-mm\\-dd\\ hh:mm:ss".equals(numFmt.getCode());
            }
        }
    }

    @Test public void testDateFmt() throws IOException {
        String fileName = "customize_date_format.xlsx";
        List<ItemFull> expectList = ItemFull.randomFull();
        new Workbook()
            .setAutoSize(true)
            .addSheet(new ListSheet<>(expectList
            , new Column("编码", "code")
            , new Column("姓名", "name")
            , new Column("日期", "date").setNumFmt("yyyy年mm月dd日 hh日mm分ss秒")
        )).writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            org.ttzero.excel.reader.Sheet sheet = reader.sheet(0);
            Iterator<org.ttzero.excel.reader.Row> iter = sheet.header(1).iterator();
            org.ttzero.excel.reader.HeaderRow header = (HeaderRow) sheet.getHeader();
            assert "编码".equals(header.get(0));
            assert "姓名".equals(header.get(1));
            assert "日期".equals(header.get(2));

            for (ItemFull expect : expectList) {
                assert iter.hasNext();
                org.ttzero.excel.reader.Row row = iter.next();
                Map<String, Object> e = row.toMap();
                assert expect.code.equals(e.get("编码"));
                assert expect.name.equals(e.get("姓名"));
                assert expect.date.getTime() / 1000 == ((Timestamp) e.get("日期")).getTime() / 1000;

                Styles styles = row.getStyles();
                int styleIndex = row.getCellStyle(2);
                NumFmt numFmt = styles.getNumFmt(styleIndex);
                assert numFmt != null && "yyyy年mm月dd日\\ hh日mm分ss秒".equals(numFmt.getCode());
            }
        }
    }

    @Test public void testNumFmt() throws IOException {
        String fileName = "customize_numfmt_full.xlsx";
        List<ItemFull> expectList = ItemFull.randomFull();
        new Workbook()
            .setAutoSize(true)
            .addSheet(new ListSheet<>(expectList
                , new Column("编码", "code")
                , new Column("姓名", "name")
                , new Column("日期", "date").setNumFmt("上午/下午hh\"時\"mm\"分\"")
                , new Column("数字", "num").setNumFmt("#,##0 ;[Red]-#,##0 ")
            )).writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            org.ttzero.excel.reader.Sheet sheet = reader.sheet(0);
            Iterator<org.ttzero.excel.reader.Row> iter = sheet.header(1).iterator();
            org.ttzero.excel.reader.HeaderRow header = (HeaderRow) sheet.getHeader();
            assert "编码".equals(header.get(0));
            assert "姓名".equals(header.get(1));
            assert "日期".equals(header.get(2));
            assert "数字".equals(header.get(3));

            for (ItemFull expect : expectList) {
                assert iter.hasNext();
                org.ttzero.excel.reader.Row row = iter.next();
                Map<String, Object> e = row.toMap();
                assert expect.code.equals(e.get("编码"));
                assert expect.name.equals(e.get("姓名"));
                assert expect.date.getTime() / 1000 == ((Timestamp) e.get("日期")).getTime() / 1000;
                assert expect.num == (Long) e.get("数字");

                Styles styles = row.getStyles();
                int styleIndex = row.getCellStyle(2);
                NumFmt numFmt = styles.getNumFmt(styleIndex);
                assert numFmt != null && "上午/下午hh\"時\"mm\"分\"".equals(numFmt.getCode());
                int styleIndex3 = row.getCellStyle(3);
                NumFmt numFmt3 = styles.getNumFmt(styleIndex3);
                assert numFmt3 != null && "#,##0\\ ;[Red]\\-#,##0\\ ".equals(numFmt3.getCode());
            }
        }
    }

    @Test public void testNegativeNumFmt() throws IOException {
        String fileName = "customize_negative.xlsx";
        List<Num> expectList;
        new Workbook()
            .setAutoSize(true)
            .addSheet(new ListSheet<>(expectList = Arrays.asList(new Num(1234565435236543436L), new Num(0), new Num(-1234565435236543436L))))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            Iterator<org.ttzero.excel.reader.Row> iter = reader.sheet(0).header(1).bind(Num.class).iterator();
            for (Num expect : expectList) {
                assert iter.hasNext();
                org.ttzero.excel.reader.Row row = iter.next();
                assert expect.equals(row.get());

                Styles styles = row.getStyles();
                int styleIndex = row.getCellStyle(0);
                NumFmt numFmt = styles.getNumFmt(styleIndex);
                assert numFmt != null && "[Blue]#,##0.00_);[Red]\\-#,##0.00_);0_)".equals(numFmt.getCode());
            }
        }
    }

    @Test public void testNumFmtWidth() {
        NumberFormat nf = NumberFormat.getInstance();
        nf.setGroupingUsed(false);
        nf.setMaximumFractionDigits(6);

        NumFmt fmt = new NumFmt("[Blue]###0.00_);[Red]-###0.00_);0_)");
        double width;

        width = fmt.calcNumWidth(nf.format(12345654352365434.36D).length());
        assert width >= 20.86D && width <= 25.63D;

        width = fmt.calcNumWidth(nf.format(-12345654352365434.36D).length());
        assert width >= 21.5D && width <= 26.63D;

        width = fmt.calcNumWidth(stringSize(1234565));
        assert width >= 11.5D && width <= 13.63D;

        width = fmt.calcNumWidth(stringSize(-1234565));
        assert width >= 12.5D && width <= 14.63D;

        fmt.setCode("[Blue]#,##0.00_);[Red]-#,##0.00_);0_)");
        width = fmt.calcNumWidth(stringSize(1234565435236543436L));
        assert width >= 29.0D && width <= 34.63D;

        width = fmt.calcNumWidth(stringSize(-1234565435236543436L));
        assert width >= 30.5D && width <= 35.63D;

        width = fmt.calcNumWidth(stringSize(1234565));
        assert width >= 13.0D && width <= 15.63D;

        width = fmt.calcNumWidth(stringSize(-1234565));
        assert width >= 14.5D && width <= 16.63D;

        fmt.setCode("[Blue]#,##0;[Red]-#,##0;0");
        width = fmt.calcNumWidth(stringSize(1234565435236543436L));
        assert width >= 25.5D && width <= 29.63D;

        width = fmt.calcNumWidth(stringSize(-1234565435236543436L));
        assert width >= 26.5D && width <= 30.63D;

        width = fmt.calcNumWidth(stringSize(1234565));
        assert width >= 9.5D && width <= 12.63D;

        width = fmt.calcNumWidth(stringSize(-1234565));
        assert width >= 10.5D && width <= 13.63D;

        fmt.setCode("yyyy-mm-dd");
        width = fmt.calcNumWidth(0);
        assert width >= 10.86D && width <= 12.63D;

        fmt.setCode("yyyy-mm-dd hh:mm:ss");
        width = fmt.calcNumWidth(0);
        assert width >= 19.86D && width <= 23.63D;

        fmt.setCode("hh:mm:ss");
        width = fmt.calcNumWidth(0);
        assert width >= 8.86D && width <= 10.63D;

        fmt.setCode("yyyy年mm月dd日 hh日mm分ss秒");
        width = fmt.calcNumWidth(0);
        assert width >= 26.86D && width <= 30.63D;
    }

    @Test public void testAutoWidth() throws IOException {
        String fileName = "Auto Width Test.xlsx";
        List<WidthTestItem> expectList = WidthTestItem.randomTestData();
        new Workbook()
            .setAutoSize(true)
            .addSheet(new ListSheet<>(expectList))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            Iterator<org.ttzero.excel.reader.Row> iter = reader.sheet(0).header(1).bind(WidthTestItem.class).iterator();
            for (WidthTestItem expect : expectList) {
                assert iter.hasNext();
                org.ttzero.excel.reader.Row row = iter.next();
                assert expect.equals(row.get());

                Styles styles = row.getStyles();
                int styleIndex = row.getCellStyle(0);
                NumFmt numFmt = styles.getNumFmt(styleIndex);
                assert numFmt != null && "#,##0_);[Red]\\-#,##0_);0_)".equals(numFmt.getCode());
                int styleIndex3 = row.getCellStyle(3);
                NumFmt numFmt3 = styles.getNumFmt(styleIndex3);
                assert numFmt3 != null && "yyyy\\-mm\\-dd\\ hh:mm:ss".equals(numFmt3.getCode());
            }
        }
    }

    @Test public void testAutoAndMaxWidth() throws IOException {
        String fileName = "Auto Max Width Test.xlsx";
        List<WidthTestItem> expectList = MaxWidthTestItem.randomTestData();
        new Workbook()
                .setAutoSize(true)
                .addSheet(new ListSheet<>(expectList))
                .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            Iterator<org.ttzero.excel.reader.Row> iter = reader.sheet(0).header(1).bind(MaxWidthTestItem.class).iterator();
            for (WidthTestItem expect : expectList) {
                assert iter.hasNext();
                org.ttzero.excel.reader.Row row = iter.next();
                assert expect.equals(row.get());

                Styles styles = row.getStyles();
                int styleIndex = row.getCellStyle(0);
                NumFmt numFmt = styles.getNumFmt(styleIndex);
                assert numFmt != null && "#,##0_);[Red]\\-#,##0_);0_)".equals(numFmt.getCode());
                int styleIndex3 = row.getCellStyle(3);
                NumFmt numFmt3 = styles.getNumFmt(styleIndex3);
                assert numFmt3 != null && "yyyy\\-mm\\-dd\\ hh:mm:ss".equals(numFmt3.getCode());
            }
        }
    }

    public static class Item {
        @ExcelColumn
        String code;
        @ExcelColumn
        String name;
        @ExcelColumn(format = "yyyy-mm-dd")
        Date date;

        public Item() { }

        static List<Item> random() {
            int n = random.nextInt(10) + 1;
            List<Item> list = new ArrayList<>(n);
            for (; n-- > 0; ) {
                Item e = new Item();
                e.code = "code" + n;
                e.name = getRandomString();
                e.date = new Timestamp(System.currentTimeMillis() - random.nextInt(9999999));
                list.add(e);
            }
            return list;
        }

        @Override
        public boolean equals(Object o) {
            if (this == o) return true;
            if (o == null || getClass() != o.getClass()) return false;
            Item item = (Item) o;
            return Objects.equals(code, item.code) &&
                Objects.equals(name, item.name) &&
                date.getTime() / 1000 == item.date.getTime() / 1000;
        }

        @Override
        public int hashCode() {
            return Objects.hash(code, name, date);
        }
    }

    public static class ItemFull extends Item {

        @ExcelColumn
        long num;

        public ItemFull() { }

        @ExcelColumn(format = "yyyy-mm-dd hh:mm:ss")
        public Date getDate() {
            return date;
        }

        static List<ItemFull> randomFull() {
            int n = random.nextInt(10) + 1;
            List<ItemFull> list = new ArrayList<>(n);
            for (; n-- > 0; ) {
                ItemFull e = new ItemFull();
                e.code = "code" + n;
                e.name = getRandomString();
                e.date = new Timestamp(System.currentTimeMillis() - random.nextInt(9999999));
                e.num = random.nextLong();
                list.add(e);
            }
            return list;
        }

        @Override
        public boolean equals(Object o) {
            if (this == o) return true;
            if (o == null || getClass() != o.getClass()) return false;
            if (!super.equals(o)) return false;
            ItemFull itemFull = (ItemFull) o;
            return num == itemFull.num;
        }

        @Override
        public int hashCode() {
            return Objects.hash(super.hashCode(), num);
        }
    }

    public static class Num {
        @ExcelColumn(format = "[Blue]#,##0.00_);[Red]-#,##0.00_);0_)")
        long num;

        public Num() { }
        Num(long num) {
            this.num = num;
        }

        @Override
        public boolean equals(Object o) {
            if (this == o) return true;
            if (o == null || getClass() != o.getClass()) return false;
            Num num1 = (Num) o;
            return num == num1.num;
        }

        @Override
        public int hashCode() {
            return Objects.hash(num);
        }
    }

    public static class WidthTestItem {
        @ExcelColumn(value = "整型", format = "#,##0_);[Red]-#,##0_);0_)")
        Integer nv;
        @ExcelColumn("字符串(en)")
        String sen;
        @ExcelColumn("字符串(中文)")
        String scn;
        @ExcelColumn(value = "日期时间", format = "yyyy-mm-dd hh:mm:ss")
        Timestamp iv;

        public WidthTestItem() { }

        public static List<WidthTestItem> randomTestData() {
            return randomTestData(WidthTestItem::new);
        }
        public static List<WidthTestItem> randomTestData(Supplier<? extends WidthTestItem> supplier) {
            int size = random.nextInt(10) + 5;
            List<WidthTestItem> list = new ArrayList<>(size);
            for (int i = 0; i < size; i++) {
                WidthTestItem o = supplier.get();
                o.nv = random.nextInt();
                o.iv = new Timestamp(System.currentTimeMillis() - random.nextInt(9999999));
                o.sen = getRandomString(20);
                o.scn = "联想笔记本电脑拯救者R7000(标压 6核 R5-5600H 16G 512G RTX3050 100%sRGB)黑";
                list.add(o);
            }
            return list;
        }

        @Override
        public boolean equals(Object o) {
            if (this == o) return true;
            if (o == null || getClass() != o.getClass()) return false;
            WidthTestItem that = (WidthTestItem) o;
            return Objects.equals(nv, that.nv) &&
                Objects.equals(sen, that.sen) &&
                Objects.equals(scn, that.scn) &&
                iv.getTime() / 1000 == that.iv.getTime() / 1000;
        }

        @Override
        public int hashCode() {
            return Objects.hash(nv, sen, scn, iv);
        }
    }

    public static class MaxWidthTestItem extends WidthTestItem {
        @ExcelColumn(value = "字符串(中文)", maxWidth = 30.86D, wrapText = true)
        public String getScn() {
            return scn;
        }

        public MaxWidthTestItem() { }
        public static List<WidthTestItem> randomTestData() {
            return randomTestData(MaxWidthTestItem::new);
        }
    }
}

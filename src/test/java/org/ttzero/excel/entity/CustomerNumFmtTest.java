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

import java.io.IOException;
import java.sql.Timestamp;
import java.text.NumberFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.List;
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
        new Workbook("customize_numfmt").setAutoSize(true).addSheet(new ListSheet<>(Item.random())).writeTo(defaultTestPath);
    }

    @Test public void testFmtInAnnotationYmdHms() throws IOException {
        new Workbook("customize_numfmt_full").setAutoSize(true).addSheet(new ListSheet<>(ItemFull.randomFull())).writeTo(defaultTestPath);
    }

    @Test public void testDateFmt() throws IOException {
        new Workbook("customize_date_format")
            .setAutoSize(true)
            .addSheet(new ListSheet<>(ItemFull.randomFull()
            , new Column("编码", "code")
            , new Column("姓名", "name")
            , new Column("日期", "date").setNumFmt("yyyy年mm月dd日 hh日mm分ss秒")
        )).writeTo(defaultTestPath);
    }

    @Test public void testNumFmt() throws IOException {
        new Workbook("customize_numfmt_full")
            .setAutoSize(true)
            .addSheet(new ListSheet<>(ItemFull.randomFull()
                , new Column("编码", "code")
                , new Column("姓名", "name")
                , new Column("日期", "date").setNumFmt("上午/下午hh\"時\"mm\"分\"")
                , new Column("数字", "num").setNumFmt("#,##0 ;[Red]-#,##0 ")
            )).writeTo(defaultTestPath);
    }

    @Test public void testNegativeNumFmt() throws IOException {
        new Workbook("customize_negative")
            .setAutoSize(true)
            .addSheet(new ListSheet<>(Arrays.asList(new Num(1234565435236543436L), new Num(0), new Num(-1234565435236543436L))))
            .writeTo(defaultTestPath);
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
        new Workbook("Auto Width Test")
            .setAutoSize(true)
            .addSheet(new ListSheet<>(WidthTestItem.randomTestData()))
            .writeTo(defaultTestPath);
    }

    @Test public void testAutoAndMaxWidth() throws IOException {
        new Workbook("Auto Max Width Test")
                .setAutoSize(true)
                .addSheet(new ListSheet<>(MaxWidthTestItem.randomTestData()))
                .writeTo(defaultTestPath);
    }

    static class Item {
        @ExcelColumn
        String code;
        @ExcelColumn
        String name;
        @ExcelColumn(format = "yyyy-mm-dd")
        Date date;

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
    }

    static class ItemFull extends Item {

        @ExcelColumn
        long num;

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
    }

    static class Num {
        @ExcelColumn(format = "[Blue]#,##0.00_);[Red]-#,##0.00_);0_)")
        long num;

        Num(long num) {
            this.num = num;
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

        public static List<WidthTestItem> randomTestData() {
            return randomTestData(WidthTestItem::new);
        }
        public static List<WidthTestItem> randomTestData(Supplier<? extends WidthTestItem> supplier) {
            int size = random.nextInt(10 + 5);
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
    }

    public static class MaxWidthTestItem extends WidthTestItem {
        @ExcelColumn(value = "字符串(中文)", maxWidth = 30.86D, wrapText = true)
        public String getScn() {
            return scn;
        }

        public static List<WidthTestItem> randomTestData() {
            return randomTestData(MaxWidthTestItem::new);
        }
    }
}

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

import java.io.IOException;
import java.sql.Timestamp;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.List;

import static org.ttzero.excel.Print.println;
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
        new Workbook("customize_numfmt").addSheet(new ListSheet<>(Item.random())).writeTo(defaultTestPath);
    }

    @Test public void testFmtInAnnotationYmdHms() throws IOException {
        new Workbook("customize_numfmt_full").addSheet(new ListSheet<>(ItemFull.randomFull())).writeTo(defaultTestPath);
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
                .addSheet(new ListSheet<>(Arrays.asList(new Num(12345678), new Num(0), new Num(-12345678))))
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
        int num;

        Num(int num) {
            this.num = num;
        }
    }
}

package org.ttzero.excel.entity;

import org.junit.Test;
import org.ttzero.excel.annotation.ExcelColumn;

import java.io.IOException;
import java.sql.Timestamp;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

public class CustomerNumFmtTest extends WorkbookTest {

    @Test public void testFmtInAnnotation() throws IOException {
        new Workbook("customize_numfmt").addSheet(new ListSheet<>(Item.random())).writeTo(defaultTestPath);
    }

    @Test public void testFmtInAnnotationYmdHms() throws IOException {
        new Workbook("customize_numfmt_full").addSheet(new ListSheet<>(ItemFull.randomFull())).writeTo(defaultTestPath);
    }

    @Test public void testDateFmt() throws IOException {
        new Workbook("customize_numfmt_full")
                .setAutoSize(true)
                .addSheet(new ListSheet<>(ItemFull.randomFull()
                , new Sheet.Column("编码", "code")
                , new Sheet.Column("姓名", "name")
                , new Sheet.Column("日期", "date").setNumFmt("yyyy年mm月dd日 hh日mm分")
        )).writeTo(defaultTestPath);
    }

    @Test public void testNumFmt() throws IOException {
        new Workbook("customize_numfmt_full")
                .setAutoSize(true)
                .addSheet(new ListSheet<>(ItemFull.randomFull()
                        , new Sheet.Column("编码", "code")
                        , new Sheet.Column("姓名", "name")
                        , new Sheet.Column("日期", "date").setNumFmt("上午/下午hh\"時\"mm\"分\"")
                        , new Sheet.Column("数字", "num").setNumFmt("#,##0")
                )).writeTo(defaultTestPath);
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
}

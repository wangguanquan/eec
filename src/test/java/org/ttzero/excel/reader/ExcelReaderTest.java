/*
 * Copyright (c) 2019, guanquan.wang@yandex.com All Rights Reserved.
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

package org.ttzero.excel.reader;

import org.junit.Test;
import org.ttzero.excel.Print;
import org.ttzero.excel.annotation.ExcelColumn;
import org.ttzero.excel.annotation.IgnoreExport;
import org.ttzero.excel.manager.ExcelType;
import org.ttzero.excel.util.DateUtil;
import org.ttzero.excel.util.FileUtil;

import java.io.IOException;
import java.net.URL;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Date;
import java.util.Iterator;

import static org.ttzero.excel.Print.println;
import static org.ttzero.excel.Print.print;
import static org.ttzero.excel.entity.WorkbookTest.getOutputTestPath;

/**
 * Create by guanquan.wang at 2019-04-26 17:42
 */
public class ExcelReaderTest {
    public static Path testResourceRoot() {
        URL url = ExcelReaderTest.class.getClassLoader().getResource(".");
        if (url == null) {
            throw new RuntimeException("Load test resources error.");
        }
        return FileUtil.isWindows()
            ? Paths.get(url.getFile().substring(1))
            : Paths.get(url.getFile());
    }

    @Test public void testReader() {
        try (ExcelReader reader = ExcelReader.read(testResourceRoot().resolve("1.xlsx"))) {
            assert reader.getType() == ExcelType.XLSX;

            AppInfo appInfo = reader.getAppInfo();
            assert "对象数组测试".equals(appInfo.getTitle());
            assert "guanquan.wang".equals(appInfo.getCreator());
            println(appInfo);

            reader.sheet(0).rows().forEach(Print::println);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    @Test public void testColumnIndex() {
        try (ExcelReader reader = ExcelReader.read(testResourceRoot().resolve("1.xlsx"))) {
            Sheet sheet = reader.sheet(0);
            for (Iterator<Row> it = sheet.iterator(); it.hasNext();) {
                Row row = it.next();
                println(row.getRowNumber()
                    + " | " + row.getFirstColumnIndex()
                    + " | " + row.getLastColumnIndex()
                    + " => " + row);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    @Test public void testReset() {
        try (ExcelReader reader = ExcelReader.read(testResourceRoot().resolve("1.xlsx"))) {

            Sheet sheet = reader.sheet(0);
            sheet.rows().forEach(Print::println);

            println("------------------");

            sheet.reset(); // Reset the row index to begging

            sheet.rows().forEach(Print::println);

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    @Test public void testForEach() {
        try (ExcelReader reader = ExcelReader.read(testResourceRoot().resolve("1.xlsx"))) {
            Sheet sheet = reader.sheet(0);

            Row header = sheet.getHeader();

            for (Iterator<Row> it = sheet.iterator(); it.hasNext(); ) {
                Row row = it.next();
                if (row.getRowNumber() == 0) continue;

                print(row.getRowNumber());
                for (int start = 0, end = row.getLastColumnIndex(); start < end; start++) {
                    print(header.getString(start));
                    print(" : ");
                    CellType type = row.getCellType(start);
                    switch (type) {
                        case DATE    : print(row.getTimestamp(start)); break;
                        case INTEGER : print(row.getInt(start))      ; break;
                        case LONG    : print(row.getLong(start))     ; break;
                        case DOUBLE  : print(row.getDouble(start))   ; break;
                        default      : print(row.getString(start))   ; break;
                    }
                    print(' ');
                }
                println();
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    @Test public void testToStandardObject() {
        try (ExcelReader reader = ExcelReader.read(testResourceRoot().resolve("1.xlsx"))) {
            reader.sheets().flatMap(Sheet::dataRows).map(row -> row.too(StandardEntry.class)).forEach(Print::println);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    @Test public void testToObject() {
        try (ExcelReader reader = ExcelReader.read(testResourceRoot().resolve("1.xlsx"))) {
            reader.sheets().flatMap(Sheet::dataRows).map(row -> row.too(Entry.class)).forEach(Print::println);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    @Test public void testToAnnotationObject() {
        try (ExcelReader reader = ExcelReader.read(testResourceRoot().resolve("1.xlsx"))) {
            reader.sheets().flatMap(Sheet::dataRows).map(row -> row.too(AnnotationEntry.class)).forEach(Print::println);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    @Test public void testToCSV() {
        try (ExcelReader reader = ExcelReader.read(testResourceRoot().resolve("1.xlsx"))) {
            reader.sheet(0).saveAsCSV(getOutputTestPath());
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static class Entry {
        @ExcelColumn("渠道ID")
        private Integer channelId;
        @ExcelColumn(value = "游戏", share = true)
        private String pro;
        @ExcelColumn
        private String account;
        @ExcelColumn("注册时间")
        private java.util.Date registered;
        @ExcelColumn("是否满30级")
        private boolean up30;
        @IgnoreExport("敏感信息不导出")
        private int id; // not export
        private String address;
        @ExcelColumn("VIP")
        private char c;

        private boolean vip;

        public boolean isUp30() {
            return up30;
        }

        /**
         * Convert game name to code
         *
         * @param pro the game nice name
         */
        public void setPro(String pro) {
            // "LOL", "WOW", "极品飞车", "守望先锋", "怪物世界"
            String code;
            switch (pro) {
                case "LOL"   : code = "1"; break;
                case "WOW"   : code = "2"; break;
                case "极品飞车": code = "3"; break;
                case "守望先锋": code = "4"; break;
                case "怪物世界": code = "5"; break;
                default: code = "0";
            }
            this.pro = code;
        }

        public void setC(char c) {
            this.c = c;
            this.vip = c == 'A';
        }

        public boolean isVip() {
            return vip;
        }

        @Override
        public String toString() {
            return channelId + " | "
                + pro + " | "
                + account + " | "
                + (registered != null ? DateUtil.toDateString(registered) : null) + " | "
                + up30 + " | "
                + c + " | "
                + isVip()
                ;
        }
    }

    public static class StandardEntry {
        private Integer channelId;
        private String pro;
        private String account;
        private java.util.Date registered;
        private boolean up30;
        private int id;
        private String address;
        private char c;

        private boolean vip;

        public void setChannelId(Integer channelId) {
            this.channelId = channelId;
        }

        public void setPro(String pro) {
            this.pro = pro;
        }

        public void setAccount(String account) {
            this.account = account;
        }

        public void setRegistered(Date registered) {
            this.registered = registered;
        }

        public void setUp30(boolean up30) {
            this.up30 = up30;
        }

        public void setId(int id) {
            this.id = id;
        }

        public void setAddress(String address) {
            this.address = address;
        }

        public void setC(char c) {
            this.c = c;
            this.vip = c == 'A';
        }

        @Override
        public String toString() {
            return channelId + " | "
                + pro + " | "
                + account + " | "
                + (registered != null ? DateUtil.toDateString(registered) : null) + " | "
                + up30 + " | "
                + c + " | "
                + vip
                ;
        }
    }

    public static class AnnotationEntry {
        private Integer channelId;
        private String pro;
        private String account;
        private java.util.Date registered;
        private boolean up30;
        private int id;
        private String address;
        private char c;

        private boolean vip;

        @ExcelColumn("渠道ID")
        public void setChannelId(Integer channelId) {
            this.channelId = channelId;
        }

        @ExcelColumn(value = "游戏")
        public void setPro(String pro) {
            this.pro = pro;
        }

        public void setAccount(String account) {
            this.account = account;
        }

        @ExcelColumn("注册时间")
        public void setRegistered(Date registered) {
            this.registered = registered;
        }

        @ExcelColumn("是否满30级")
        public void setUp30(boolean up30) {
            this.up30 = up30;
        }

        public void setId(int id) {
            this.id = id;
        }

        public void setAddress(String address) {
            this.address = address;
        }

        @ExcelColumn("VIP")
        public void setC(char c) {
            this.c = c;
            this.vip = c == 'A';
        }

        @Override
        public String toString() {
            return channelId + " | "
                + pro + " | "
                + account + " | "
                + (registered != null ? DateUtil.toDateString(registered) : null) + " | "
                + up30 + " | "
                + c + " | "
                + vip
                ;
        }
    }
}

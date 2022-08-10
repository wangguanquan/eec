/*
 * Copyright (c) 2017-2019, guanquan.wang@yandex.com All Rights Reserved.
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

import org.junit.Ignore;
import org.junit.Test;
import org.ttzero.excel.Print;
import org.ttzero.excel.annotation.ExcelColumn;
import org.ttzero.excel.annotation.IgnoreExport;
import org.ttzero.excel.annotation.IgnoreImport;
import org.ttzero.excel.annotation.RowNum;
import org.ttzero.excel.entity.ListObjectSheetTest;
import org.ttzero.excel.entity.WorkbookTest;
import org.ttzero.excel.util.DateUtil;
import org.ttzero.excel.util.FileUtil;
import org.ttzero.excel.util.StringUtil;

import java.io.File;
import java.io.IOException;
import java.math.BigDecimal;
import java.net.URL;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.util.Arrays;
import java.util.Date;
import java.util.Iterator;
import java.util.List;
import java.util.Objects;
import java.util.stream.Collectors;

import static org.ttzero.excel.Print.println;
import static org.ttzero.excel.Print.print;
import static org.ttzero.excel.entity.Sheet.int2Col;
import static org.ttzero.excel.entity.WorkbookTest.getOutputTestPath;
import static org.ttzero.excel.reader.ExcelReader.COPY_ON_MERGED;
import static org.ttzero.excel.reader.ExcelReader.VALUE_AND_CALC;
import static org.ttzero.excel.reader.ExcelReader.VALUE_ONLY;
import static org.ttzero.excel.reader.ExcelReader.cellRangeToLong;
import static org.ttzero.excel.util.StringUtil.swap;

/**
 * @author guanquan.wang at 2019-04-26 17:42
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

    @Test public void testReader() throws IOException {
        File[] files = testResourceRoot().toFile().listFiles((dir, name) -> name.endsWith(".xlsx"));
        if (files != null) {
            for (File file : files) {
                testReader(file.toPath());
            }
        }
    }

    @Test public void testMergedReader() throws IOException {
        File[] files = testResourceRoot().toFile().listFiles((dir, name) -> name.endsWith(".xlsx"));
        if (files != null) {
            for (File file : files) {
                testReader(file.toPath(), COPY_ON_MERGED);
            }
        }
    }

    @Test public void testFormulaReader() throws IOException {
        File[] files = testResourceRoot().toFile().listFiles((dir, name) -> name.endsWith(".xlsx"));
        if (files != null) {
            for (File file : files) {
                testFormulaReader(file.toPath());
            }
        }
    }

    @Test public void testColumnIndex() throws IOException {
        try (ExcelReader reader = ExcelReader.read(testResourceRoot().resolve("1.xlsx"))) {
            Sheet sheet = reader.sheet(0);
            for (Iterator<Row> it = sheet.iterator(); it.hasNext();) {
                Row row = it.next();
                println(row.getRowNum()
                    + " | " + row.getFirstColumnIndex()
                    + " | " + row.getLastColumnIndex()
                    + " => " + row);
            }
        }
    }

    @Test public void testReset() throws IOException {
        try (ExcelReader reader = ExcelReader.read(testResourceRoot().resolve("1.xlsx"))) {

            Sheet sheet = reader.sheet(0);
            sheet.rows().forEach(Print::println);

            println("------------------");

            sheet.reset(); // Reset the row index to begging

            sheet.rows().forEach(Print::println);

        }
    }

    @Test public void testForEach() throws IOException {
        try (ExcelReader reader = ExcelReader.read(testResourceRoot().resolve("1.xlsx"))) {
            Sheet sheet = reader.sheet(0);

            Row header = sheet.getHeader();

            for (Iterator<Row> it = sheet.iterator(); it.hasNext(); ) {
                Row row = it.next();
                if (row.getRowNum() == 0) continue;

                print(row.getRowNum());
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
        }
    }

    @Test public void testToStandardObject() throws IOException {
        try (ExcelReader reader = ExcelReader.read(testResourceRoot().resolve("1.xlsx"))) {
            reader.sheets().flatMap(Sheet::dataRows).map(row -> row.too(StandardEntry.class)).forEach(Print::println);
        }
    }

    @Test public void testToObject() throws IOException {
        try (ExcelReader reader = ExcelReader.read(testResourceRoot().resolve("1.xlsx"))) {
            reader.sheets().flatMap(Sheet::dataRows).map(row -> row.too(Entry.class)).forEach(Print::println);
        }
    }

    @Test public void testToAnnotationObject() throws IOException {
        try (ExcelReader reader = ExcelReader.read(testResourceRoot().resolve("1.xlsx"))) {
            reader.sheets().flatMap(Sheet::dataRows).map(row -> row.too(AnnotationEntry.class)).forEach(Print::println);
        }
    }

    @Test public void testReaderByName() throws IOException {
        try (ExcelReader reader = ExcelReader.read(testResourceRoot().resolve("1.xlsx"))) {
            reader.sheet(0).dataIterator().forEachRemaining(row -> {
                print(row.getInt("渠道ID")); print(" | ");
                print(row.getString("游戏")); print(" | ");
                print(row.getString("account")); print(" | ");
                print(row.getDate("注册时间")); print(" | ");
                print(row.getBoolean("是否满30级")); print(" | ");
                print(row.getChar("VIP"));
                println();
            });
        }
    }

    @Test public void testFilter() throws IOException {
        try (ExcelReader reader = ExcelReader.read(testResourceRoot().resolve("1.xlsx"))) {
            String[] games = reader.sheet(0)
                .dataRows()
                .map(row -> row.getString("游戏"))
                .distinct()
                .toArray(String[]::new);
            print(Arrays.toString(games));
        }
    }

    @Test public void testToCSV() throws IOException {
        try (ExcelReader reader = ExcelReader.read(testResourceRoot().resolve("1.xlsx"))) {
            reader.sheet(0).saveAsCSV(getOutputTestPath());
        }
    }

    @Test public void test_81() throws IOException {
        try (ExcelReader reader = ExcelReader.read(testResourceRoot().resolve("#81.xlsx"))) {
            List<Customer> list = reader.sheets().flatMap(Sheet::dataRows).map(row -> row.to(Customer.class)).collect(Collectors.toList());

            for (Customer c : list) System.out.println(c);
        }
    }

    @Test public void testDimension() throws IOException {
        try (ExcelReader reader = ExcelReader.read(testResourceRoot().resolve("#81.xlsx"))) {
            Dimension dimension = reader.sheet(0).getDimension();
            System.out.println(dimension);

            assert dimension.firstRow == 1;
            assert dimension.lastRow == 6;
            assert dimension.firstColumn == 1;
            assert dimension.lastColumn == 2;
        }
    }

    @Test public void testDimensionConstructor() {
        Dimension dimension = Dimension.of("A1:C5");
        assert "A1:C5".equals(dimension.toString());

        assert dimension.firstRow == 1;
        assert dimension.firstColumn == 1;
        assert dimension.lastRow == 5;
        assert dimension.lastColumn == 3;
    }

    @Test public void testFormula() throws IOException {
        try (ExcelReader reader = ExcelReader.read(testResourceRoot().resolve("formula.xlsx"))) {
            // Read value
            reader.sheets().flatMap(Sheet::rows).forEach(Print::println);

            if (reader.hasFormula()) {
                // Reset and parse formula
                reader.parseFormula().sheets().flatMap(sheet -> {
                    println("----------------" + sheet.getName() + "----------------");
                    return sheet.dataRows();
                }).forEach(row -> {
                    for (int i = row.fc; i < row.lc; i++) {
                        if (row.hasFormula(i)) {
                            print(int2Col(i + 1));
                            print(row.getRowNum());
                            print("=");
                            print(row.getFormula(i));
                            println();
                        }
                    }
                });
            }
        }
    }


    @Test public void testSearch() {
        long[] array = { 131075L, 327683L };
        int column = 2, row = 3;
        boolean h = Arrays.binarySearch(array, ((column + 1) & 0x7FFF) | ((long) row) << 16) >= 0;
        println(h);

        print(cellRangeToLong("AA10"));
    }

    @Test public void testClassBind() throws IOException {
        try (ExcelReader reader = ExcelReader.read(testResourceRoot().resolve("1.xlsx"))) {
            reader.sheet(0).bind(Entry.class).dataRows().forEach(row -> {
                // Use bind...get...
                // Getting and convert to specify Entry
                Entry entry = row.get();
                System.out.println(entry.toString());
            });
        }
    }

    @Test public void testClassSharedBind() throws IOException {
        try (ExcelReader reader = ExcelReader.read(testResourceRoot().resolve("1.xlsx"))) {
            reader.sheet(0).bind(Entry.class).dataRows().forEach(row -> {
                // Use bind...geet...
                // Getting and convert to specify Entry, the entry is shared in memory
                Entry entry = row.geet();
                System.out.println(entry.toString());
            });
        }
    }

    @Test public void testHeaderString() throws IOException {
        try (ExcelReader reader = ExcelReader.read(testResourceRoot().resolve("1.xlsx"))) {
            reader.sheets().flatMap(sheet -> {
                println("----------------" + sheet.getName() + "----------------");
                println(sheet.getHeader());
                return sheet.dataRows();
            }).forEach(Print::println);
        }
    }

    @Test public void testHeaderString2() throws IOException {
        try (ExcelReader reader = ExcelReader.read(testResourceRoot().resolve("1.xlsx"))) {
            reader.sheets().flatMap(sheet -> {
                println("----------------" + sheet.getName() + "----------------");
                sheet.bind(Entry.class);
                println(sheet.getHeader());
                return sheet.dataRows();
            }).forEach(row -> println((Entry) row.get()));
        }
    }

    @Test public void testVersionFilter() {
        char[] chars = "..0...3...7.SNAPSHOT.".toCharArray();
        int i = 0;
        for (int j = 0; j < chars.length; j++) {
            if (chars[j] >= '0' && chars[j] <= '9' || chars[j] == '.' && i > 0 && chars[i - 1] != '.')
                chars[i++] = chars[j];
        }
        String version = i > 0 ? new String(chars, 0, chars[i - 1] != '.' ? i : i - 1) : "1.0.0";
        assert "0.3.7".equals(version);
    }

    @Test public void testSort() {
        int index = 6;
        String[] values = {"ref", "B2:B8", "t", "shared","si", "0"};
        // Sort like t, si, ref
        for (int i = 0, len = index >> 1; i < len; i++) {
            int _i = i << 1;
            int vl = values[_i].length();
            if (vl - 1 == i) {
                continue;
            }
            // Will be sort
            int _n = vl - 1;
            swap(values, _n << 1, _i);
            swap(values, (_n << 1) + 1, _i + 1);
        }

        assert "t".equals(values[0]);
        assert "shared".equals(values[1]);
        assert "si".equals(values[2]);
        assert "0".equals(values[3]);
        assert "ref".equals(values[4]);
        assert "B2:B8".equals(values[5]);
    }

    @Test public void testMergeFunc() throws IOException {
        try (ExcelReader reader = ExcelReader.read(testResourceRoot().resolve("formula.xlsx"))) {
//            reader.copyOnMergeCells().sheets().flatMap(s -> {
//                println("----------------" + s.getName() + "----------------");
//                println("dimension: " + s.getDimension());
//                return s.rows();
//            }).forEach(Print::println);
            reader.copyOnMergeCells().sheets().flatMap(Sheet::rows).forEach(Print::println);
        }
    }


    @Test public void testMergeExcel() throws IOException {
        try (ExcelReader reader = ExcelReader.read(testResourceRoot().resolve("merge.xlsx"))) {
            reader.parseFormula().sheets().flatMap(Sheet::rows).forEach(Print::println);
        }
    }

    @Test public void testMergeExcel2() throws IOException {
        try (ExcelReader reader = ExcelReader.read(testResourceRoot().resolve("#150.xlsx"))) {
            reader.sheets().flatMap(Sheet::rows).forEach(Print::println);
        }
    }

    @Ignore
    @Test public void testReaderLarge() throws IOException {
        try (ExcelReader reader = ExcelReader.read(WorkbookTest.getOutputTestPath().resolve("large07.xlsx"))) {
            long n = reader.sheets().flatMap(Sheet::dataRows).map(row -> row.too(LargeData.class)).count();
            assert n == 500000;
        }
    }

    @Test public void testAllType() throws IOException {
        try (ExcelReader reader = ExcelReader.read(testResourceRoot().resolve("all type.xlsx"))) {
            reader.sheets().flatMap(Sheet::dataRows)
                .map(row -> row.too(ListObjectSheetTest.AllType.class))
                .forEach(Print::println);
        }
    }

    @Test public void testBoxAllType() throws IOException {
        try (ExcelReader reader = ExcelReader.read(testResourceRoot().resolve("all type.xlsx"))) {
            reader.sheets().flatMap(Sheet::dataRows)
                .map(row -> row.too(ListObjectSheetTest.BoxAllType.class))
                .forEach(Print::println);
        }
    }

    @Test public void testNumber2ExcelFormula() throws IOException {
        testFormulaReader(testResourceRoot().resolve("Number2Excel.xlsx"));
    }

    @Test public void testResetToEntry() throws IOException {
        try (ExcelReader reader = ExcelReader.read(testResourceRoot().resolve("1.xlsx"))) {
            println("--------0--------");
            reader.sheet(0).reset().rows().forEach(Print::println);

            println("--------1--------");
            reader.sheet(0).dataRows().map(row -> row.too(Entry.class)).forEach(Print::println);

            println("--------2--------");
            reader.sheet(0).reset().dataRows().map(row -> row.too(Entry.class)).forEach(Print::println);

            Sheet sheet = reader.sheet(0);
            println("--------3--------");
            sheet.reset().dataRows().map(row -> row.too(Entry.class)).forEach(Print::println);

            println("--------4--------");
            sheet.reset().rows().forEach(Print::println);

            println("--------5--------");
            reader.sheet(0).reset().rows().forEach(Print::println);

            println("--------6--------");
            reader.sheet(0).asCalcSheet().reset().rows().forEach(Print::println);

            println("--------7--------");
            reader.sheet(0).asCalcSheet().reset().rows().forEach(Print::println);

            println("--------8--------");
            reader.sheet(0).asMergeSheet().reset().rows().forEach(Print::println);

            println("--------9--------");
            reader.sheet(0).reset().rows().forEach(Print::println);
        }
    }

    private void testReader(Path path) throws IOException {
        testReader(path, VALUE_ONLY);
    }

    private void testReader(Path path, int option) throws IOException {
        println("----------" + path.getFileName() + "----------");
        try (ExcelReader reader = ExcelReader.read(path, option)) {
            println(reader.getAppInfo());

            reader.sheets()
                .peek(sheet -> println("--------" + sheet.getName() + "--------" + sheet.getDimension()))
                .flatMap(Sheet::rows)
                .forEach(row -> println(row.getRowNum() + "|: " + row.toString()));

        }
    }


    private void testFormulaReader(Path path) throws IOException {
        println("----------" + path.getFileName() + "----------");
        try (ExcelReader reader = ExcelReader.read(path, VALUE_AND_CALC)) {
            // Read formula
            reader.sheets().flatMap(Sheet::rows).forEach(row -> {
                for (int i = row.fc; i < row.lc; i++) {
                    if (row.hasFormula(i)) {
                        print(new Dimension(row.getRowNum(), (short) (i + 1)));
                        print(": ");
                        print(row.getFormula(i));
                        println('|');
                    }
                }
            });
        }
    }

    @Test public void testToObject2() throws IOException {
        try (ExcelReader reader = ExcelReader.read(testResourceRoot().resolve("test-fixed-row.xlsx"))) {
            reader.sheet(0).rows().forEach(row -> {
                if (row.getRowNum() == 1) {
                    assert "我是固定表头".equals(row.getString(0));
                } else if (row.getRowNum() == 2) {
                    assert "我是内容".equals(row.getString(0));
                }
            });
        }
    }

    @Test public void testReadEmptyCell() throws IOException {
        try (ExcelReader reader = ExcelReader.read(testResourceRoot().resolve("#169.xlsx"))) {
            reader.sheets().peek(sheet -> println(sheet.getName() + ": " + sheet.getDimension())).flatMap(Sheet::rows).forEach(Print::println);
            reader.sheets().peek(sheet -> {
                sheet.reset();
                println(sheet.getName() + ": " + sheet.getDimension());
            }).flatMap(Sheet::rows).forEach(Print::println);
        }
    }

    @Test public void testReadDrawings() throws IOException {
        try (ExcelReader reader = ExcelReader.read(testResourceRoot().resolve("drawing.xlsx"))) {
            reader.sheets().peek(sheet -> println(sheet.getName() + ": " + sheet.getDimension())).flatMap(Sheet::rows).forEach(Print::println);

            // From workbook`
            List<Drawings.Picture> pictures = reader.listPictures();
            assert pictures.size() == 5;
            pictures.forEach(Print::println);

            // Copy images
            for (Drawings.Picture pic : pictures) {
                Path dest = Paths.get("target/excel/drawing/", pic.sheet.getName(), pic.localPath.getFileName().toString());
                if (!Files.exists(dest.getParent())) FileUtil.mkdir(dest.getParent());
                Files.copy(pic.localPath, dest, StandardCopyOption.REPLACE_EXISTING);
                assert Files.size(pic.localPath) == Files.size(dest);
            }

            // From worksheet
            reader.sheets().forEach(sheet -> {
                List<Drawings.Picture> pictures1 = sheet.listPictures();
                if (sheet.getName().equals("Sheet1")) {
                    assert pictures1.size() == 4;
                } else assert !sheet.getName().equals("Sheet2") || pictures1.size() == 1;
                pictures1.forEach(Print::println);
            });
        }
    }

    @Test public void test175() throws IOException {
        try (ExcelReader reader = ExcelReader.read(testResourceRoot().resolve("#175.xlsx"))) {
            reader.sheet(0).rows().filter(row -> row.getRowNum() > 7 && !row.isEmpty()).forEach(row -> println(row.getDouble(4)));
            reader.sheet(0).reset();
            println("-------------DECIMAL-------------");
            reader.sheet(0).rows().filter(row -> row.getRowNum() > 7 && !row.isEmpty()).forEach(row -> println(row.getDecimal(4)));
            reader.sheet(0).reset();
            reader.sheet(0).rows()
                    .filter(row -> row.getRowNum() > 6 && !row.isEmpty())
                    .map(row -> row.to(O.class))
                    .filter(Objects::nonNull)
                    .forEach(Print::println);
        }
    }

    @Test public void test226() throws IOException {
        final String[] arr = {"ab", "", "r", "y", "", "6", "nrge"};
        try (ExcelReader reader = ExcelReader.read(testResourceRoot().resolve("#226.xlsx"))) {
            String[] array = reader.sheet(0).rows().map(row -> row.getString(0)).toArray(String[]::new);
            assert Arrays.equals(arr, array);
        }
    }

    @Test public void test354() throws IOException {
        try (ExcelReader reader = ExcelReader.read(testResourceRoot().resolve("#175.xlsx"))) {
            OO[] list = reader.sheet(0).rows()
                    .filter(row -> row.getRowNum() > 6 && !row.isEmpty())
                    .map(row -> row.to(OO.class))
                    .filter(Objects::nonNull)
                    .toArray(OO[]::new);

            assert "rowNum: 8 => fbaNo: FBA15DRV4JP4U000001, refId: 2Z91JHMR, price: 0.08, weight: 0.070000000000000007, brand: TEYASI, productName: 手机充电头".equals(list[0].toString());
            assert "rowNum: 9 => fbaNo: FBA15DRV4JP4U000002, refId: 2Z91JHMR, price: 0.08, weight: 0.070000000000000007, brand: TEYASI, productName: 手机充电头".equals(list[1].toString());
        }
    }

    @Test public void testLargeMerge() throws IOException {
        try (ExcelReader reader = ExcelReader.read(testResourceRoot().resolve("largeMerged.xlsx"))) {
            Grid grid = reader.sheet(0).asMergeSheet().getMergeGrid();
            assert grid.test(3, 1);
            assert grid.test(382, 1);
            assert grid.test(722, 2);
            assert grid.test(1374, 2);
            assert grid.test(2101, 10);
            assert grid.test(2201, 6);
            assert !grid.test(2113, 5);
        }
    }

    public static class O {
        @ExcelColumn("亚马逊FBA子单号/箱唛号")
        private String fbaNo;

        @ExcelColumn("Reference ID（亚马逊追踪编码）")
        private String refId;

        @ExcelColumn("单个产品申报单价")
        private BigDecimal price;

        @ExcelColumn("单个产品净重KG(必填)")
        private BigDecimal weight;

        @Override
        public String toString() {
            return "fbaNo: " + fbaNo + ", refId: " + refId + ", price: " + price + ", weight: " + weight;
        }
    }


    public static class Customer {
        @ExcelColumn("客户编码")
        private String code;
        @ExcelColumn("人员工号")
        private String name;

        public String getCode() {
            return code;
        }

        public void setCode(String code) {
            this.code = code;
        }

        public String getName() {
            return name;
        }

        public void setName(String name) {
            this.name = name;
        }

        @Override
        public String toString() {
            return code + ": " + name;
        }
    }

    public static class Entry {
        @RowNum
        private int num;
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
            return num + " | " + channelId + " | "
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
        private int rowNum;

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

        @RowNum
        public void setRowNum(int rowNum) {
            this.rowNum = rowNum;
        }

        @Override
        public String toString() {
            return rowNum + " | " + channelId + " | "
                + pro + " | "
                + account + " | "
                + (registered != null ? DateUtil.toDateString(registered) : null) + " | "
                + up30 + " | "
                + c + " | "
                + vip
                ;
        }
    }

    public static class LargeData {
        private String str1;
        private String str2;
        private String str3;
        private String str4;
        private String str5;
        private String str6;
        private String str7;
        private String str8;
        private String str9;
        private String str10;
        private String str11;
        private String str12;
        private String str13;
        private String str14;
        private String str15;
        private String str16;
        private String str17;
        private String str18;
        private String str19;
        private String str20;
        private String str21;
        private String str22;
        private String str23;
        private String str24;
        private String str25;

        public String getStr1() {
            return str1;
        }

        public void setStr1(String str1) {
            this.str1 = str1;
        }

        public String getStr2() {
            return str2;
        }

        public void setStr2(String str2) {
            this.str2 = str2;
        }

        public String getStr3() {
            return str3;
        }

        public void setStr3(String str3) {
            this.str3 = str3;
        }

        public String getStr4() {
            return str4;
        }

        public void setStr4(String str4) {
            this.str4 = str4;
        }

        public String getStr5() {
            return str5;
        }

        public void setStr5(String str5) {
            this.str5 = str5;
        }

        public String getStr6() {
            return str6;
        }

        public void setStr6(String str6) {
            this.str6 = str6;
        }

        public String getStr7() {
            return str7;
        }

        public void setStr7(String str7) {
            this.str7 = str7;
        }

        public String getStr8() {
            return str8;
        }

        public void setStr8(String str8) {
            this.str8 = str8;
        }

        public String getStr9() {
            return str9;
        }

        public void setStr9(String str9) {
            this.str9 = str9;
        }

        public String getStr10() {
            return str10;
        }

        public void setStr10(String str10) {
            this.str10 = str10;
        }

        public String getStr11() {
            return str11;
        }

        public void setStr11(String str11) {
            this.str11 = str11;
        }

        public String getStr12() {
            return str12;
        }

        public void setStr12(String str12) {
            this.str12 = str12;
        }

        public String getStr13() {
            return str13;
        }

        public void setStr13(String str13) {
            this.str13 = str13;
        }

        public String getStr14() {
            return str14;
        }

        public void setStr14(String str14) {
            this.str14 = str14;
        }

        public String getStr15() {
            return str15;
        }

        public void setStr15(String str15) {
            this.str15 = str15;
        }

        public String getStr16() {
            return str16;
        }

        public void setStr16(String str16) {
            this.str16 = str16;
        }

        public String getStr17() {
            return str17;
        }

        public void setStr17(String str17) {
            this.str17 = str17;
        }

        public String getStr18() {
            return str18;
        }

        public void setStr18(String str18) {
            this.str18 = str18;
        }

        public String getStr19() {
            return str19;
        }

        public void setStr19(String str19) {
            this.str19 = str19;
        }

        public String getStr20() {
            return str20;
        }

        public void setStr20(String str20) {
            this.str20 = str20;
        }

        public String getStr21() {
            return str21;
        }

        public void setStr21(String str21) {
            this.str21 = str21;
        }

        public String getStr22() {
            return str22;
        }

        public void setStr22(String str22) {
            this.str22 = str22;
        }

        public String getStr23() {
            return str23;
        }

        public void setStr23(String str23) {
            this.str23 = str23;
        }

        public String getStr24() {
            return str24;
        }

        public void setStr24(String str24) {
            this.str24 = str24;
        }

        public String getStr25() {
            return str25;
        }

        public void setStr25(String str25) {
            this.str25 = str25;
        }
    }

    public static class Goods {
        @ExcelColumn("商品编码")
        private String no;
        @ExcelColumn("商品名称")
        private String name;
        @ExcelColumn("*品牌")
        private String brand;
        @ExcelColumn("*订货号")
        private String buyNo;
        @ExcelColumn("型号")
        private String model;
        @ExcelColumn("*单位")
        private String unit;
        @ExcelColumn("税率（不填默认为0）")
        private BigDecimal taxRate;
        @ExcelColumn("*含税单价（元）")
        private BigDecimal price;
        @ExcelColumn("*采购数量")
        private BigDecimal count;

        public String getNo() {
            return no;
        }

        public void setNo(String no) {
            this.no = no;
        }

        public String getName() {
            return name;
        }

        public void setName(String name) {
            this.name = name;
        }

        public String getBrand() {
            return brand;
        }

        public void setBrand(String brand) {
            this.brand = brand;
        }

        public String getBuyNo() {
            return buyNo;
        }

        public void setBuyNo(String buyNo) {
            this.buyNo = buyNo;
        }

        public String getModel() {
            return model;
        }

        public void setModel(String model) {
            this.model = model;
        }

        public String getUnit() {
            return unit;
        }

        public void setUnit(String unit) {
            this.unit = unit;
        }

        public BigDecimal getTaxRate() {
            return taxRate;
        }

        public void setTaxRate(BigDecimal taxRate) {
            this.taxRate = taxRate;
        }

        public BigDecimal getPrice() {
            return price;
        }

        public void setPrice(BigDecimal price) {
            this.price = price;
        }

        public BigDecimal getCount() {
            return count;
        }

        public void setCount(BigDecimal count) {
            this.count = count;
        }

        @Override
        public String toString() {
            return buyNo + " " + price + " " + count;
        }
    }

    public static class OO {
        @IgnoreImport
        @ExcelColumn(colIndex = 3)
        private BigDecimal price;

        @ExcelColumn(colIndex = 1)
        private String refId;

        private BigDecimal weight;

        private String brandName, productName;

        @ExcelColumn(colIndex = 0)
        private String fbaNo;

        @RowNum
        private Integer rowNum;

        @Override
        public String toString() {
            return "rowNum: " + rowNum + " => fbaNo: " + fbaNo + ", refId: " + refId + ", price: " + price + ", weight: " + weight + ", brand: " + brandName + ", productName: " + productName;
        }

        @ExcelColumn("单个产品净重KG(必填)")
        public void abc(BigDecimal weight) {
            this.weight = weight;
        }

        @ExcelColumn(colIndex = 3)
        public void setPriceString(String price) {
            if (StringUtil.isNotEmpty(price)) {
                try {
                    this.price = new BigDecimal(price);
                } catch (Exception e) {
                    // Ignore
                }
            }
        }

        @ExcelColumn(colIndex = 5)
        public void setBrandName(String brandName) {
            this.brandName = brandName;
        }

        @ExcelColumn(colIndex = 2)
        public void setName(String productName) {
            this.productName = productName;
        }
    }

}

/*
 * Copyright (c) 2017-2019, guanquan.wang@hotmail.com All Rights Reserved.
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

package org.ttzero.excel.entity.csv;

import org.junit.Before;
import org.junit.Ignore;
import org.junit.Test;
import org.ttzero.excel.entity.CSVSheet;
import org.ttzero.excel.entity.Workbook;
import org.ttzero.excel.entity.WorkbookTest;
import org.ttzero.excel.reader.ExcelReader;
import org.ttzero.excel.reader.Row;
import org.ttzero.excel.reader.Sheet;
import org.ttzero.excel.util.CSVUtil;
import org.ttzero.excel.util.StringUtil;

import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Iterator;
import java.util.List;
import java.util.concurrent.atomic.AtomicInteger;

import static org.junit.Assert.assertArrayEquals;
import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertTrue;
import static org.ttzero.excel.util.FileUtil.exists;

/**
 * @author guanquan.wang at 2019-09-26 10:07
 */
public class CsvToExcelTest extends WorkbookTest {
    private Path path;
    @Before public void before() throws IOException {
        path = getOutputTestPath().resolve("1.csv");

        // Create a test file
        if (!exists(path)) {
            int column = random.nextInt(10) + 1, row = random.nextInt(100) + 1;

            // storage column type
            // 0 string
            // 1 char
            // 2 int
            // 3 float
            // 4 double
            int[] types = new int[column];
            for (int i = 0; i < column; i++) {
                types[i] = random.nextInt(5);
            }

            try (CSVUtil.Writer writer = CSVUtil.newWriter(path)) {
                for (int i = 0; i < row; i++) {
                    for (int c = 0; c < column; c++) {
                        switch (types[c]) {
                            case 0:
                                writer.write(getRandomString());
                                break;
                            case 1:
                                writer.write(charArray[random.nextInt(charArray.length)]);
                                break;
                            case 2:
                                writer.write(random.nextInt());
                                break;
                            case 3:
                                writer.write(random.nextFloat());
                                break;
                            case 4:
                                writer.write(random.nextDouble());
                                break;
                        }
                    }
                    // break row
                    writer.newLine();
                }
            }
        }
    }

    @Test public void testFromPath() throws IOException {
        String fileName = "csv path test.xlsx";
        new Workbook()
            .addSheet(new CSVSheet(path))
            .writeTo(getOutputTestPath().resolve(fileName));

        List<String[]> expectList = CSVUtil.read(path);
        try (ExcelReader reader = ExcelReader.read(getOutputTestPath().resolve(fileName))) {
            Iterator<org.ttzero.excel.reader.Row> iter = reader.sheet(0).iterator();
            for (String[] expect : expectList) {
                assertTrue(iter.hasNext());
                org.ttzero.excel.reader.Row row = iter.next();
                for (int i = 0; i < expect.length; i++) {
                    if (expect[i] != null) {
                        assertEquals(expect[i], row.getString(i));
                    } else {
                        assertTrue(StringUtil.isEmpty(row.getString(i)));
                    }
                }
            }
        }
    }

    @Test public void testFromInputStream() throws IOException {
        String fileName = "csv inputstream test.xlsx";
        new Workbook()
            .addSheet(new CSVSheet(Files.newInputStream(path)))
            .writeTo(getOutputTestPath().resolve(fileName));

        List<String[]> expectList = CSVUtil.read(path);
        try (ExcelReader reader = ExcelReader.read(getOutputTestPath().resolve(fileName))) {
            Iterator<org.ttzero.excel.reader.Row> iter = reader.sheet(0).iterator();
            for (String[] expect : expectList) {
                assertTrue(iter.hasNext());
                org.ttzero.excel.reader.Row row = iter.next();
                for (int i = 0; i < expect.length; i++) {
                    if (expect[i] != null) {
                        assertEquals(expect[i], row.getString(i));
                    } else {
                        assertTrue(StringUtil.isEmpty(row.getString(i)));
                    }
                }
            }
        }
    }

    @Test public void testFromReader() throws IOException {
        String fileName = "csv reader test.xlsx";
        new Workbook()
            .addSheet(new CSVSheet(Files.newBufferedReader(path)))
            .writeTo(getOutputTestPath().resolve(fileName));

        List<String[]> expectList = CSVUtil.read(path);
        try (ExcelReader reader = ExcelReader.read(getOutputTestPath().resolve(fileName))) {
            Iterator<org.ttzero.excel.reader.Row> iter = reader.sheet(0).iterator();
            for (String[] expect : expectList) {
                assertTrue(iter.hasNext());
                org.ttzero.excel.reader.Row row = iter.next();
                for (int i = 0; i < expect.length; i++) {
                    if (expect[i] != null) {
                        assertEquals(expect[i], row.getString(i));
                    } else {
                        assertTrue(StringUtil.isEmpty(row.getString(i)));
                    }
                }
            }
        }
    }

    @Test public void testHasHeaderFromPath() throws IOException {
        String fileName = "csv header path test.xlsx";
        new Workbook()
            .addSheet(new CSVSheet(path))
            .writeTo(getOutputTestPath().resolve(fileName));

        List<String[]> expectList = CSVUtil.read(path);
        try (ExcelReader reader = ExcelReader.read(getOutputTestPath().resolve(fileName))) {
            Iterator<org.ttzero.excel.reader.Row> iter = reader.sheet(0).iterator();
            for (String[] expect : expectList) {
                assertTrue(iter.hasNext());
                org.ttzero.excel.reader.Row row = iter.next();
                for (int i = 0; i < expect.length; i++) {
                    if (expect[i] != null) {
                        assertEquals(expect[i], row.getString(i));
                    } else {
                        assertTrue(StringUtil.isEmpty(row.getString(i)));
                    }
                }
            }
        }
    }

    @Test public void testWriterCharsetGBK() throws IOException {
        final String fileName = "write-with-gbk.xlsx";
        String[] expectList = {"中文", "123"};
        Path distPath = getOutputTestPath().resolve("write-with-gbk.csv");
        try (CSVUtil.Writer writer = CSVUtil.newWriter(distPath, Charset.forName("GBK"))) {
            for (String v : expectList) {
                writer.write(v);
            }
        }

        try (CSVUtil.Reader reader = CSVUtil.newReader(distPath, Charset.forName("GBK"))) {
            CSVUtil.RowsIterator iter = reader.iterator();
            assertTrue(iter.hasNext());
            String[] readList = iter.next();
            assertArrayEquals(expectList, readList);
        }

        // CSV to Excel
        new Workbook().addSheet(new CSVSheet(distPath).setCharset(Charset.forName("GBK")).ignoreHeader()).writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            Iterator<Row> iter = reader.sheet(0).rows().iterator();
            assertTrue(iter.hasNext());
            Row row = iter.next();
            assertEquals(row.getString(0), expectList[0]);
            assertEquals(row.getString(1), expectList[1]);
            assertFalse(iter.hasNext());
        }
    }

    @Test public void testUTF8BOM() throws IOException {
        final String fileName = "write-with-utf8-bom.xlsx";
        String[] expectList = {"中文", "123"};
        Path distPath = getOutputTestPath().resolve("write-with-utf8-bom.csv");
        try (CSVUtil.Writer writer = CSVUtil.newWriter(distPath, StandardCharsets.UTF_8).writeWithBom()) {
            for (String v : expectList) {
                writer.write(v);
            }
        }

        try (InputStream is = Files.newInputStream(distPath)) {
            byte[] bytes = new byte[3];
            int n = is.read(bytes);
            assertEquals(3, n);
            assertArrayEquals(bytes, new byte[] {(byte) 239, (byte) 187, (byte) 191});
        }

        try (CSVUtil.Reader reader = CSVUtil.newReader(distPath)) {
            CSVUtil.RowsIterator iter = reader.iterator();
            assertTrue(iter.hasNext());
            String[] readList = iter.next();
            assertArrayEquals(expectList, readList);
        }

        // CSV to Excel
        new Workbook().addSheet(new CSVSheet(distPath).ignoreHeader()).writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            Iterator<Row> iter = reader.sheet(0).rows().iterator();
            assertTrue(iter.hasNext());
            Row row = iter.next();
            assertEquals(row.getString(0), expectList[0]);
            assertEquals(row.getString(1), expectList[1]);
            assertFalse(iter.hasNext());
        }
    }

    @Test public void testUTF16BEBOM() throws IOException {
        final String fileName = "write-with-utf16BE-bom.xlsx";
        String[] expectList = {"中文", "123"};
        Path distPath = getOutputTestPath().resolve("write-with-utf16BE-bom.csv");
        try (CSVUtil.Writer writer = CSVUtil.newWriter(distPath, StandardCharsets.UTF_16BE).writeWithBom()) {
            for (String v : expectList) {
                writer.write(v);
            }
        }

        try (InputStream is = Files.newInputStream(distPath)) {
            byte[] bytes = new byte[2];
            int n = is.read(bytes);
            assertEquals(2, n);
            assertArrayEquals(bytes, new byte[] {-2, -1});
        }

        try (CSVUtil.Reader reader = CSVUtil.newReader(distPath)) {
            CSVUtil.RowsIterator iter = reader.iterator();
            assertTrue(iter.hasNext());
            String[] readList = iter.next();
            assertArrayEquals(expectList, readList);
        }

        // CSV to Excel
        new Workbook().addSheet(new CSVSheet(distPath).setCharset(StandardCharsets.UTF_16BE).ignoreHeader()).writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            Iterator<Row> iter = reader.sheet(0).rows().iterator();
            assertTrue(iter.hasNext());
            Row row = iter.next();
            assertEquals(row.getString(0), expectList[0]);
            assertEquals(row.getString(1), expectList[1]);
            assertFalse(iter.hasNext());
        }
    }

    @Ignore
    @Test public void testIah94s() throws IOException {
        final String fileName = "3343494.xlsx";
        try (CSVUtil.Writer writer = CSVUtil.newWriter(defaultTestPath.resolve("3343494.csv"))) {
            writer.write("ID");
            writer.write("NAME");
            writer.newLine();
            StringBuilder buf = new StringBuilder("ab");
            for (int i = 0; i < 3343494; i++) {
                writer.write(i);
                buf.append(i);
                writer.write(buf.toString());
                writer.newLine();
                buf.delete(2, buf.length());
            }
        }

        new Workbook()
                .addSheet(new CSVSheet(defaultTestPath.resolve("3343494.csv")))
                .writeTo(defaultTestPath.resolve(fileName));

        AtomicInteger oi = new AtomicInteger(0);
        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            boolean noneMatch = reader.sheets().flatMap(Sheet::dataRows).noneMatch(row -> {
                int i = oi.getAndIncrement();
                return row.getInt(0) == i && ("ab" + i).equals(row.getString(1));
            });
            assertFalse(noneMatch);
        }
    }

    @Test public void testFromInputStreamSpecifyDelimiter() throws IOException {
        String fileName = "csv inputstream specify delimiter test.xlsx";
        new Workbook()
            .addSheet(new CSVSheet(Files.newInputStream(path)).setDelimiter(','))
            .writeTo(getOutputTestPath().resolve(fileName));

        List<String[]> expectList = CSVUtil.read(path, ',');
        try (ExcelReader reader = ExcelReader.read(getOutputTestPath().resolve(fileName))) {
            Iterator<org.ttzero.excel.reader.Row> iter = reader.sheet(0).iterator();
            for (String[] expect : expectList) {
                assertTrue(iter.hasNext());
                org.ttzero.excel.reader.Row row = iter.next();
                for (int i = 0; i < expect.length; i++) {
                    if (expect[i] != null) {
                        assertEquals(expect[i], row.getString(i));
                    } else {
                        assertTrue(StringUtil.isEmpty(row.getString(i)));
                    }
                }
            }
        }
    }
}

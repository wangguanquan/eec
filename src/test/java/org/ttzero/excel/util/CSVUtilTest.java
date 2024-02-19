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

package org.ttzero.excel.util;

import org.junit.Before;
import org.junit.Ignore;
import org.junit.Test;

import java.io.IOException;
import java.nio.charset.Charset;
import java.nio.file.Path;
import java.util.Arrays;
import java.util.List;

import static org.junit.Assert.assertArrayEquals;
import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;
import static org.ttzero.excel.Print.println;
import static org.ttzero.excel.entity.WorkbookTest.charArray;
import static org.ttzero.excel.entity.WorkbookTest.defaultTestPath;
import static org.ttzero.excel.entity.WorkbookTest.getRandomString;
import static org.ttzero.excel.entity.WorkbookTest.random;

/**
 * @author guanquan.wang at 2019-02-14 17:02
 */
public class CSVUtilTest {

    Path path = defaultTestPath.resolve("1.csv");
    @Before public void before() throws IOException {
        // Create a test file
        testWriter();
    }

    @Test public void testReader() throws IOException {
        List<String[]> rows = CSVUtil.read(path);
        rows.forEach(t -> println(Arrays.toString(t)));
    }

    @Test public void testStream() throws IOException {
        try (CSVUtil.Reader reader = CSVUtil.newReader(path)) {
            reader.stream().forEach(t -> println(Arrays.toString(t)));
        }
    }

    @Test public void testStreamShare() throws IOException {
        try (CSVUtil.Reader reader = CSVUtil.newReader(path)) {
            reader.sharedStream().forEach(t -> println(Arrays.toString(t)));
        }
    }

    @Test public void testIterator() throws IOException {
        try (CSVUtil.RowsIterator iterator = CSVUtil.newReader(path).iterator()) {
            while (iterator.hasNext()) {
                String[] rows = iterator.next();
                println(Arrays.toString(rows));
            }
        }
    }

    @Test public void testSharedIterator() throws IOException {
        try (CSVUtil.RowsIterator iterator = CSVUtil.newReader(path).sharedIterator()) {
            while (iterator.hasNext()) {
                String[] rows = iterator.next();
                println(Arrays.toString(rows));
            }
        }
    }

    @Test public void testWriteBoolean() throws IOException {
        try (CSVUtil.Writer writer = CSVUtil.newWriter(path)) {
            writer.write(true);
            writer.write(false);
        }
        List<String[]> strings = CSVUtil.read(path);
        assertEquals(strings.size(), 1);
        assertEquals(strings.get(0).length, 2);
        assertEquals(String.valueOf(true).toUpperCase(), strings.get(0)[0]);
        assertEquals(String.valueOf(false).toUpperCase(), strings.get(0)[1]);
    }

    @Ignore @Test public void testWriteChar() throws IOException {
        char c = charArray[random.nextInt(charArray.length)];
        try (CSVUtil.Writer writer = CSVUtil.newWriter(path)) {
            writer.writeChar(c);
        }

        List<String[]> strings = CSVUtil.read(path, ',');
        assertEquals(strings.size(), 1);
        assertEquals(strings.get(0).length, 1);
        // will be trim
        if (c == '\n' || c == '\t') {
            assertTrue(strings.get(0)[0].isEmpty());
        } else {
            assertEquals(String.valueOf(c), strings.get(0)[0]);
        }
    }

    @Test public void testWriteInt() throws IOException {
        int n1 = random.nextInt();
        int n2 = random.nextInt(1024);
        int zero = 0;
        int min = Integer.MIN_VALUE;
        int max = Integer.MAX_VALUE;
        int min1 = min + 1;
        int max1 = max - 1;
        try (CSVUtil.Writer writer = CSVUtil.newWriter(path)) {
            writer.write(n1);
            writer.write(n2);
            writer.write(zero);
            writer.write(min);
            writer.write(max);
            writer.write(min1);
            writer.write(max1);
        }

        List<String[]> strings = CSVUtil.read(path);
        assertEquals(strings.size(), 1);
        assertEquals(strings.get(0).length, 7);
        String[] ss = strings.get(0);
        assertEquals(String.valueOf(n1), ss[0]);
        assertEquals(String.valueOf(n2), ss[1]);
        assertEquals(String.valueOf(zero), ss[2]);
        assertEquals(String.valueOf(min), ss[3]);
        assertEquals(String.valueOf(max), ss[4]);
        assertEquals(String.valueOf(min1), ss[5]);
        assertEquals(String.valueOf(max1), ss[6]);
    }

    @Test public void testWriteLong() throws IOException {
        long l1 = random.nextLong();
        long l2 = random.nextLong();
        long l3 = random.nextLong();
        long l4 = random.nextLong();
        long min = Long.MIN_VALUE;
        long max = Long.MAX_VALUE;
        long min1 = min + 1;
        long max1 = max - 1;
        try (CSVUtil.Writer writer = CSVUtil.newWriter(path)) {
            writer.write(l1);
            writer.write(l2);
            writer.newLine(); // new line
            writer.write(l3);
            writer.write(l4);
            writer.newLine(); // new line
            writer.write(min);
            writer.write(max);
            writer.newLine(); // new line
            writer.write(min1);
            writer.write(max1);
        }

        List<String[]> strings = CSVUtil.read(path);
        assertEquals(strings.size(), 4);
        assertEquals(strings.get(0).length, 2);
        assertEquals(strings.get(1).length, 2);
        assertEquals(strings.get(2).length, 2);
        assertEquals(strings.get(3).length, 2);
        assertEquals(String.valueOf(l1), strings.get(0)[0]);
        assertEquals(String.valueOf(l2), strings.get(0)[1]);
        assertEquals(String.valueOf(l3), strings.get(1)[0]);
        assertEquals(String.valueOf(l4), strings.get(1)[1]);
        assertEquals(String.valueOf(min), strings.get(2)[0]);
        assertEquals(String.valueOf(max), strings.get(2)[1]);
        assertEquals(String.valueOf(min1), strings.get(3)[0]);
        assertEquals(String.valueOf(max1), strings.get(3)[1]);
    }

    @Test public void testWriteFloat() throws IOException {
        float f1 = random.nextFloat();
        float f2 = random.nextFloat();
        float f3 = random.nextFloat();
        try (CSVUtil.Writer writer = CSVUtil.newWriter(path)) {
            writer.write(f1);
            writer.newLine(); // new line
            writer.write(f2);
            writer.newLine(); // new line
            writer.write(f3);
        }

        List<String[]> strings = CSVUtil.read(path);
        assertEquals(strings.size(), 3);
        assertEquals(strings.get(0).length, 1);
        assertEquals(strings.get(1).length, 1);
        assertEquals(strings.get(2).length, 1);
        assertEquals(String.valueOf(f1), strings.get(0)[0]);
        assertEquals(String.valueOf(f2), strings.get(1)[0]);
        assertEquals(String.valueOf(f3), strings.get(2)[0]);
    }

    @Test public void testWriteDouble() throws IOException{
        double d1 = random.nextDouble();
        double d2 = random.nextDouble();
        double d3 = random.nextDouble();
        try (CSVUtil.Writer writer = CSVUtil.newWriter(path)) {
            writer.write(d1);
            writer.write(d2);
            writer.write(d3);
        }

        List<String[]> strings = CSVUtil.read(path);
        assertEquals(strings.size(), 1);
        assertEquals(strings.get(0).length, 3);
        assertEquals(String.valueOf(d1), strings.get(0)[0]);
        assertEquals(String.valueOf(d2), strings.get(0)[1]);
        assertEquals(String.valueOf(d3), strings.get(0)[2]);
    }

    @Test public void testWriteString() throws IOException {
        int n = random.nextInt(10) + 1;
        String[] src = new String[n];
        for (int i = 0; i < n; i++) {
            src[i] = getRandomString();
        }
        try (CSVUtil.Writer writer = CSVUtil.newWriter(path)) {
            for (String s : src) {
                writer.write(s);
            }
        }

        List<String[]> strings = CSVUtil.read(path);
        assertEquals(strings.size(), 1);
        assertEquals(strings.get(0).length, n);

        assertArrayEquals(src, strings.get(0));
    }

    @Test public void testEndWithLF() throws IOException {
        int n = random.nextInt(10) + 1;
        String[] src = new String[n];
        for (int i = 0; i < n; i++) {
            src[i] = getRandomString();
        }
        try (CSVUtil.Writer writer = CSVUtil.newWriter(path)) {
            for (String s : src) {
                writer.write(s);
            }
            writer.newLine(); // end with LF
        }

        List<String[]> strings = CSVUtil.read(path);
        assertEquals(strings.size(), 1);
        assertEquals(strings.get(0).length, n);
        assertArrayEquals(src, strings.get(0));
    }

    @Test public void testLineEndWithComma() throws IOException {
        int n = random.nextInt(10) + 1;
        String[] src = new String[n];
        for (int i = 0; i < n; i++) {
            src[i] = getRandomString();
        }
        char comma = ',';
        try (CSVUtil.Writer writer = CSVUtil.newWriter(path, comma)) {
            for (String s : src) {
                writer.write(s);
                writer.writeEmpty(); // comma is the last character
                writer.newLine(); // new line
            }
        }

        List<String[]> strings = CSVUtil.read(path);
        assertEquals(strings.size(), n);
        for (int i = 0; i < n; i++) {
            assertEquals(strings.get(i).length, 2);
            assertEquals(src[i], strings.get(i)[0]);
            assertEquals("", strings.get(i)[1]);
        }
    }

    @Test public void testLF() throws IOException {
        String s1 = "abc,12\njiuh"; // LF
        String s2 = "42\r\n843432545\"fjid";  // CR LF
        try (CSVUtil.Writer writer = CSVUtil.newWriter(path)) {
            writer.write(123);
            writer.write("中文");
            writer.write(s1);
            writer.write(s2);
            writer.write(true);
            writer.write("fjdskfhinsfainwepcfjskldafdnjsh fdslaf jdsk djsk djska hdsuiafgsdkafhdsuiafhsuiafusfhdsa");
            writer.write("fsa");
        }

        List<String[]> strings = CSVUtil.read(path);
        assertEquals(strings.size(), 1);
        assertEquals(strings.get(0).length, 7);
        assertEquals(s1, strings.get(0)[2]);
        assertEquals(s2, strings.get(0)[3]);
    }

    @Test public void testGBKCharset() throws IOException {
        String s1 = "双引号或单引号中间的一切都是字符串";
        String s2 = "中文测试使用GBK";

        Charset GBK = Charset.forName("GBK");

        try (CSVUtil.Writer writer = CSVUtil.newWriter(path, GBK)) {
            writer.write(s1);
            writer.write(s2);
        }

        // Use UTF-8
        try {
            CSVUtil.read(path);
        } catch (IOException e) { // A charset.MalformedInputException occur
            assertTrue(true);
        }

        // Use GBK
        List<String[]> strings = CSVUtil.read(path, GBK);
        assertEquals(strings.size(), 1);
        assertEquals(strings.get(0).length, 2);
        assertEquals(s1, strings.get(0)[0]);
        assertEquals(s2, strings.get(0)[1]);


        // Reset charset
        try (CSVUtil.Writer writer = CSVUtil.newWriter(path)) {
            writer.write(s1);
            writer.write(s2);
        }
    }

    @Test public void testEuropeanComma() throws IOException {
        // A comma and the value separator is a semicolon
        char comma = ';';
        int n = random.nextInt(10) + 10;
        String[] src = new String[n];
        for (int i = 0; i < n; i++) {
            src[i] = getRandomString();
        }
        try (CSVUtil.Writer writer = CSVUtil.newWriter(path, comma)) {
            for (String s : src) {
                writer.write(s);
            }
        }

        List<String[]> strings = CSVUtil.read(path, comma);
        assertEquals(strings.size(), 1);
        assertArrayEquals(src, strings.get(0));

        strings = CSVUtil.read(path);
        assertEquals(strings.size(), 1);
        assertArrayEquals(src, strings.get(0));
    }

    public void testWriter() throws IOException {
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

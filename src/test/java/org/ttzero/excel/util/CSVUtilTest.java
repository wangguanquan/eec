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
            for (; iterator.hasNext(); ) {
                String[] rows = iterator.next();
                println(Arrays.toString(rows));
            }
        }
    }

    @Test public void testSharedIterator() throws IOException {
        try (CSVUtil.RowsIterator iterator = CSVUtil.newReader(path).sharedIterator()) {
            for (; iterator.hasNext(); ) {
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
        assert strings.size() == 1;
        assert strings.get(0).length == 2;
        assert String.valueOf(true).toUpperCase().equals(strings.get(0)[0]);
        assert String.valueOf(false).toUpperCase().equals(strings.get(0)[1]);
    }

    @Ignore @Test public void testWriteChar() throws IOException {
        char c = charArray[random.nextInt(charArray.length)];
        try (CSVUtil.Writer writer = CSVUtil.newWriter(path)) {
            writer.writeChar(c);
        }

        List<String[]> strings = CSVUtil.read(path, ',');
        assert strings.size() == 1;
        assert strings.get(0).length == 1;
        // will be trim
        if (c == '\n' || c == '\t') {
            assert strings.get(0)[0].isEmpty();
        } else {
            assert String.valueOf(c).equals(strings.get(0)[0]);
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
        assert strings.size() == 1;
        assert strings.get(0).length == 7;
        String[] ss = strings.get(0);
        assert String.valueOf(n1).equals(ss[0]);
        assert String.valueOf(n2).equals(ss[1]);
        assert String.valueOf(zero).equals(ss[2]);
        assert String.valueOf(min).equals(ss[3]);
        assert String.valueOf(max).equals(ss[4]);
        assert String.valueOf(min1).equals(ss[5]);
        assert String.valueOf(max1).equals(ss[6]);
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
        assert strings.size() == 4;
        assert strings.get(0).length == 2;
        assert strings.get(1).length == 2;
        assert strings.get(2).length == 2;
        assert strings.get(3).length == 2;
        assert String.valueOf(l1).equals(strings.get(0)[0]);
        assert String.valueOf(l2).equals(strings.get(0)[1]);
        assert String.valueOf(l3).equals(strings.get(1)[0]);
        assert String.valueOf(l4).equals(strings.get(1)[1]);
        assert String.valueOf(min).equals(strings.get(2)[0]);
        assert String.valueOf(max).equals(strings.get(2)[1]);
        assert String.valueOf(min1).equals(strings.get(3)[0]);
        assert String.valueOf(max1).equals(strings.get(3)[1]);
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
        assert strings.size() == 3;
        assert strings.get(0).length == 1;
        assert strings.get(1).length == 1;
        assert strings.get(2).length == 1;
        assert String.valueOf(f1).equals(strings.get(0)[0]);
        assert String.valueOf(f2).equals(strings.get(1)[0]);
        assert String.valueOf(f3).equals(strings.get(2)[0]);
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
        assert strings.size() == 1;
        assert strings.get(0).length == 3;
        assert String.valueOf(d1).equals(strings.get(0)[0]);
        assert String.valueOf(d2).equals(strings.get(0)[1]);
        assert String.valueOf(d3).equals(strings.get(0)[2]);
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
        assert strings.size() == 1;
        assert strings.get(0).length == n;

        assert Arrays.toString(src).equals(Arrays.toString(strings.get(0)));
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
        assert strings.size() == 1;
        assert strings.get(0).length == n;
        assert Arrays.toString(src).equals(Arrays.toString(strings.get(0)));
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
        assert strings.size() == n;
        for (int i = 0; i < n; i++) {
            assert strings.get(i).length == 2;
            assert src[i].equals(strings.get(i)[0]);
            assert "".equals(strings.get(i)[1]);
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
        assert strings.size() == 1;
        assert strings.get(0).length == 7;
        assert s1.equals(strings.get(0)[2]);
        assert s2.equals(strings.get(0)[3]);
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
            assert true;
        }

        // Use GBK
        List<String[]> strings = CSVUtil.read(path, GBK);
        assert strings.size() == 1;
        assert strings.get(0).length == 2;
        assert s1.equals(strings.get(0)[0]);
        assert s2.equals(strings.get(0)[1]);


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
        println(Arrays.toString(src));
        try (CSVUtil.Writer writer = CSVUtil.newWriter(path, comma)) {
            for (String s : src) {
                writer.write(s);
            }
        }

        List<String[]> strings = CSVUtil.read(path, comma);
        assert strings.size() == 1;
        println(Arrays.toString(strings.get(0)));
        assert Arrays.toString(src).equals(Arrays.toString(strings.get(0)));

        strings = CSVUtil.read(path);
        assert strings.size() == 1;
        println(Arrays.toString(strings.get(0)));
        assert Arrays.toString(src).equals(Arrays.toString(strings.get(0)));
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

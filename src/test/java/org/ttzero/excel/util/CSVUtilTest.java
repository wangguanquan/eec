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
import java.net.URL;
import java.nio.charset.Charset;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Arrays;
import java.util.List;
import java.util.Random;

import static org.ttzero.excel.Print.println;
import static org.ttzero.excel.util.FileUtil.exists;
import static org.ttzero.excel.util.FileUtil.isWindows;

/**
 * @author guanquan.wang at 2019-02-14 17:02
 */
public class CSVUtilTest {
    private Path path;
    private Random random;
    private final char[] base = "abcdefghijklmnopqrstuvwxyz,ABCDEFGHIJKLMNOPQRSTUVWXYZ\t0123456789\"".toCharArray();
    private final char[][] cache_char_array = new char[25][];


    @Before public void before() {
        for (int i = 0; i < cache_char_array.length; i++) {
            cache_char_array[i] = new char[i + 1];
        }
        random = new Random();
        URL url = CSVUtilTest.class.getClassLoader().getResource(".");
        if (url == null) {
            throw new RuntimeException("Load test resources error.");
        }
        path = isWindows() ? Paths.get(url.getFile().substring(1)) : Paths.get(url.getFile());
        path = path.resolve("1.csv");

        // Create a test file
        if (!exists(path)) {
            testWriter();
        }
    }

    @Test public void testReader() {
        try {
            List<String[]> rows = CSVUtil.read(path);
            rows.forEach(t -> println(Arrays.toString(t)));
        } catch (IOException e) {
            e.printStackTrace();
            assert false;
        }
    }

    @Test public void testStream() {
        try (CSVUtil.Reader reader = CSVUtil.newReader(path)) {
            reader.stream()
                .forEach(t -> println(Arrays.toString(t)));
        } catch (IOException e) {
            e.printStackTrace();
            assert false;
        }
    }

    @Test public void testStreamShare() {
        try (CSVUtil.Reader reader = CSVUtil.newReader(path)) {
            reader.sharedStream()
                .forEach(t -> println(Arrays.toString(t)));
        } catch (IOException e) {
            e.printStackTrace();
            assert false;
        }
    }

    @Test public void testIterator() {
        try (CSVUtil.RowsIterator iterator = CSVUtil.newReader(path).iterator()) {
            for (; iterator.hasNext(); ) {
                String[] rows = iterator.next();
                println(Arrays.toString(rows));
            }
        } catch (IOException e) {
            e.printStackTrace();
            assert false;
        }
    }

    @Test public void testSharedIterator() {
        try (CSVUtil.RowsIterator iterator = CSVUtil.newReader(path).sharedIterator()) {
            for (; iterator.hasNext(); ) {
                String[] rows = iterator.next();
                println(Arrays.toString(rows));
            }
        } catch (IOException e) {
            e.printStackTrace();
            assert false;
        }
    }

    @Test public void testWriteBoolean() {
        try (CSVUtil.Writer writer = CSVUtil.newWriter(path)) {
            writer.write(true);
            writer.write(false);
        } catch (IOException e) {
            e.printStackTrace();
            assert false;
        }
        try {
            List<String[]> strings = CSVUtil.read(path);
            assert strings.size() == 1;
            assert strings.get(0).length == 2;
            assert String.valueOf(true).toUpperCase().equals(strings.get(0)[0]);
            assert String.valueOf(false).toUpperCase().equals(strings.get(0)[1]);
        } catch (IOException e) {
            e.printStackTrace();
            assert false;
        }
    }

    @Ignore
    @Test public void testWriteChar() {
        char c = base[random.nextInt(base.length)];
        try (CSVUtil.Writer writer = CSVUtil.newWriter(path)) {
            writer.writeChar(c);
        } catch (IOException e) {
            e.printStackTrace();
            assert false;
        }

        try {
            List<String[]> strings = CSVUtil.read(path, ',');
            assert strings.size() == 1;
            assert strings.get(0).length == 1;
            // will be trim
            if (c == '\n' || c == '\t') {
                assert strings.get(0)[0].isEmpty();
            } else {
                assert String.valueOf(c).equals(strings.get(0)[0]);
            }
        } catch (IOException e) {
            e.printStackTrace();
            assert false;
        }
    }

    @Test public void testWriteInt() {
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
        } catch (IOException e) {
            e.printStackTrace();
            assert false;
        }
        try {
            List<String[]> strings = CSVUtil.read(path);
            assert strings.size() == 1;
            assert strings.get(0).length == 7;
            assert String.valueOf(n1).equals(strings.get(0)[0]);
            assert String.valueOf(n2).equals(strings.get(0)[1]);
            assert String.valueOf(zero).equals(strings.get(0)[2]);
            assert String.valueOf(min).equals(strings.get(0)[3]);
            assert String.valueOf(max).equals(strings.get(0)[4]);
            assert String.valueOf(min1).equals(strings.get(0)[5]);
            assert String.valueOf(max1).equals(strings.get(0)[6]);
        } catch (IOException e) {
            e.printStackTrace();
            assert false;
        }
    }

    @Test public void testWriteLong() {
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
        } catch (IOException e) {
            e.printStackTrace();
            assert false;
        }
        try {
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
        } catch (IOException e) {
            e.printStackTrace();
            assert false;
        }

    }

    @Test public void testWriteFloat() {
        float f1 = random.nextFloat();
        float f2 = random.nextFloat();
        float f3 = random.nextFloat();
        try (CSVUtil.Writer writer = CSVUtil.newWriter(path)) {
            writer.write(f1);
            writer.newLine(); // new line
            writer.write(f2);
            writer.newLine(); // new line
            writer.write(f3);
        } catch (IOException e) {
            e.printStackTrace();
            assert false;
        }
        try {
            List<String[]> strings = CSVUtil.read(path);
            assert strings.size() == 3;
            assert strings.get(0).length == 1;
            assert strings.get(1).length == 1;
            assert strings.get(2).length == 1;
            assert String.valueOf(f1).equals(strings.get(0)[0]);
            assert String.valueOf(f2).equals(strings.get(1)[0]);
            assert String.valueOf(f3).equals(strings.get(2)[0]);
        } catch (IOException e) {
            e.printStackTrace();
            assert false;
        }
    }

    @Test public void testWriteDouble() {
        double d1 = random.nextDouble();
        double d2 = random.nextDouble();
        double d3 = random.nextDouble();
        try (CSVUtil.Writer writer = CSVUtil.newWriter(path)) {
            writer.write(d1);
            writer.write(d2);
            writer.write(d3);
        } catch (IOException e) {
            e.printStackTrace();
            assert false;
        }
        try {
            List<String[]> strings = CSVUtil.read(path);
            assert strings.size() == 1;
            assert strings.get(0).length == 3;
            assert String.valueOf(d1).equals(strings.get(0)[0]);
            assert String.valueOf(d2).equals(strings.get(0)[1]);
            assert String.valueOf(d3).equals(strings.get(0)[2]);
        } catch (IOException e) {
            e.printStackTrace();
            assert false;
        }
    }

    @Test public void testWriteString() {
        int n = random.nextInt(10) + 1;
        String[] src = new String[n];
        for (int i = 0; i < n; i++) {
            src[i] = randomString();
        }
        try (CSVUtil.Writer writer = CSVUtil.newWriter(path)) {
            for (String s : src) {
                writer.write(s);
            }
        } catch (IOException e) {
            e.printStackTrace();
            assert false;
        }
        try {
            List<String[]> strings = CSVUtil.read(path);
            assert strings.size() == 1;
            assert strings.get(0).length == n;

            assert Arrays.toString(src).equals(Arrays.toString(strings.get(0)));
        } catch (IOException e) {
            e.printStackTrace();
            assert false;
        }
    }

    @Test public void testEndWithLF() {
        int n = random.nextInt(10) + 1;
        String[] src = new String[n];
        for (int i = 0; i < n; i++) {
            src[i] = randomString();
        }
        try (CSVUtil.Writer writer = CSVUtil.newWriter(path)) {
            for (String s : src) {
                writer.write(s);
            }
            writer.newLine(); // end with LF
        } catch (IOException e) {
            e.printStackTrace();
            assert false;
        }

        try {
            List<String[]> strings = CSVUtil.read(path);
            assert strings.size() == 1;
            assert strings.get(0).length == n;
            assert Arrays.toString(src).equals(Arrays.toString(strings.get(0)));
        } catch (IOException e) {
            e.printStackTrace();
            assert false;
        }
    }

    @Test public void testLineEndWithComma() {
        int n = random.nextInt(10) + 1;
        String[] src = new String[n];
        for (int i = 0; i < n; i++) {
            src[i] = randomString();
        }
        char comma = ',';
        try (CSVUtil.Writer writer = CSVUtil.newWriter(path, comma)) {
            for (String s : src) {
                writer.write(s);
                writer.writeEmpty(); // comma is the last character
                writer.newLine(); // new line
            }
        } catch (IOException e) {
            e.printStackTrace();
            assert false;
        }

        try {
            List<String[]> strings = CSVUtil.read(path);
            assert strings.size() == n;
            for (int i = 0; i < n; i++) {
                assert strings.get(i).length == 2;
                assert src[i].equals(strings.get(i)[0]);
                assert "".equals(strings.get(i)[1]);
            }
        } catch (IOException e) {
            e.printStackTrace();
            assert false;
        }
    }

    @Test public void testLF() {
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
        } catch (IOException e) {
            e.printStackTrace();
            assert false;
        }

        try {
            List<String[]> strings = CSVUtil.read(path);
            assert strings.size() == 1;
            assert strings.get(0).length == 7;
            assert s1.equals(strings.get(0)[2]);
            assert s2.equals(strings.get(0)[3]);
        } catch (IOException e) {
            e.printStackTrace();
            assert false;
        }
    }

    @Test public void testGBKCharset() {
        String s1 = "双引号或单引号中间的一切都是字符串";
        String s2 = "中文测试使用GBK";

        Charset GBK = Charset.forName("GBK");

        try (CSVUtil.Writer writer = CSVUtil.newWriter(path, GBK)) {
            writer.write(s1);
            writer.write(s2);
        } catch (IOException e) {
            e.printStackTrace();
            assert false;
        }

        // Use UTF-8
        try {
            CSVUtil.read(path);
        } catch (IOException e) { // A charset.MalformedInputException occur
            assert true;
        }

        // Use GBK
        try {
            List<String[]> strings = CSVUtil.read(path, GBK);
            assert strings.size() == 1;
            assert strings.get(0).length == 2;
            assert s1.equals(strings.get(0)[0]);
            assert s2.equals(strings.get(0)[1]);
        } catch (IOException e) {
            e.printStackTrace();
            assert false;
        }


        // Reset charset
        try (CSVUtil.Writer writer = CSVUtil.newWriter(path)) {
            writer.write(s1);
            writer.write(s2);
        } catch (IOException e) {
            e.printStackTrace();
            assert false;
        }
    }

    @Test public void testEuropeanComma() {
        // A comma and the value separator is a semicolon
        char comma = ';';
        int n = random.nextInt(10) + 10;
        String[] src = new String[n];
        for (int i = 0; i < n; i++) {
            src[i] = randomString();
        }
        println(Arrays.toString(src));
        try (CSVUtil.Writer writer = CSVUtil.newWriter(path, comma)) {
            for (String s : src) {
                writer.write(s);
            }
        } catch (IOException e) {
            e.printStackTrace();
            assert false;
        }

        try {
            List<String[]> strings = CSVUtil.read(path, comma);
            assert strings.size() == 1;
            println(Arrays.toString(strings.get(0)));
            assert Arrays.toString(src).equals(Arrays.toString(strings.get(0)));
        } catch (IOException e) {
            e.printStackTrace();
            assert false;
        }

        try {
            List<String[]> strings = CSVUtil.read(path);
            assert strings.size() == 1;
            println(Arrays.toString(strings.get(0)));
            assert Arrays.toString(src).equals(Arrays.toString(strings.get(0)));
        } catch (IOException e) {
            e.printStackTrace();
            assert false;
        }
    }

    @Test public void testWriter() {
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
                            writer.write(randomString());
                            break;
                        case 1:
                            writer.write(base[random.nextInt(base.length)]);
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
        } catch (IOException e) {
            e.printStackTrace();
            assert false;
        }
    }

    private String randomString() {
        int len = random.nextInt(cache_char_array.length - 1);
        if (len < 5) len = 5;
        char[] cache = cache_char_array[len - 1];
        for (int j = 0; j < len; j++) {
            cache[j] = base[random.nextInt(base.length)];
        }
        return new String(cache);
    }
}

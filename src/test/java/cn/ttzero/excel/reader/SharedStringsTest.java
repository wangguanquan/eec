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

package cn.ttzero.excel.reader;

import org.junit.Test;

import java.io.IOException;
import java.nio.file.Path;
import java.nio.file.Paths;

import static cn.ttzero.excel.Print.println;
import static cn.ttzero.excel.Print.print;
import static cn.ttzero.excel.entity.WorkbookTest.getRandomString;

/**
 * Create by guanquan.wang at 2019-05-09 21:16
 */
public class SharedStringsTest {
    @Test public void test1() {
        for (int i = 1, n; ; i++) {
            if ((n = i << 6) > 0) {
//                println(n);
                assert i * 64 == i << 6;
            } else break;
        }
    }

    @Test public void test2() {
        for (int i = 1; i <= 20; i++) {
            print(i << 6);
            print(' ');
            println(Integer.toBinaryString(i << 6));
        }
    }

    @Test public void test3() {
        int n = 0x7FFFFFFF >> 6 << 6;
        for (int i = 0; i < 0x7FFFFFFF; i++) {
            assert (i & n) != i || i % 64 == 0;
        }
    }

    @Test public void test4() {
        for (int i = 128; i <= 192; i++) {
            println(i >> 6);
        }
    }

    @Test public void test5() {
        try (SharedStrings.IndexSharedStringTable sst = new SharedStrings.IndexSharedStringTable()) {
            sst.push('a');
            sst.push('b');

            String value;
            value = sst.get(0);
            println(value);
            assert value.equals("a");

            value = sst.get(1);
            println(value);
            assert value.equals("b");

            String[] array = new String[2];
            int n = sst.batch(0, array);
            for (int i = 0; i < n; i++) {
                println(array[i]);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    @Test public void test6() {
        try (SharedStrings.IndexSharedStringTable sst = new SharedStrings.IndexSharedStringTable()) {
            int length = 1000;
            String[] buf = new String[length];
            for (int i = 0; i < length; i++)
                buf[i] = getRandomString();

            for (String s : buf) {
                sst.push(s);
            }

//            for (int i = 0; i < length; i++) {
//                assert sst.get(i).equals(buf[i]);
//            }

            int fromIndex = 436, size = 1000;
            String[] _buf = new String[size];
            size = sst.batch(fromIndex, _buf);
            assert size == length - fromIndex;
            for (int i = 0; i < size; i++) {
                assert _buf[i].equals(buf[fromIndex + i]);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    @Test public void test7() throws IOException {
        SharedStrings.IndexSharedStringTable sst = new SharedStrings.IndexSharedStringTable();
        long start = System.currentTimeMillis();
        for (int i = 0; i < 10_000_000; i++) {
            sst.push(getRandomString());
        }
        System.out.println(System.currentTimeMillis() - start);

        start = System.currentTimeMillis();

        for (int i = 0; i < 10_000_000; i++) {
            sst.get(i);
        }
        System.out.println(System.currentTimeMillis() - start);
    }

    @Test public void test8() throws IOException {
        Path path = Paths.get("C:\\Users\\wangguanquan\\AppData\\Local\\Temp\\+8633905961752560043.sst.idx");
        SharedStrings.IndexSharedStringTable sst = new SharedStrings.IndexSharedStringTable(path);
        long start = System.currentTimeMillis();
        int length = 512, n = 0;
        String[] array = new String[length];
        for (int i = 0; i < 10_000_000; ) {
            int size = sst.batch(i, array);
            i += size;
            n++;
        }
        System.out.println(n);
        System.out.println(System.currentTimeMillis() - start);
    }
}

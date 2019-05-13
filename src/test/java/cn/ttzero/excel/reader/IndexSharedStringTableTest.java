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
import static cn.ttzero.excel.entity.WorkbookTest.getRandomString;

/**
 * Create by guanquan.wang at 2019-05-10 21:42
 */
public class IndexSharedStringTableTest {

    @Test public void test1() {
        try (IndexSharedStringTable sst = new IndexSharedStringTable()) {
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

    @Test public void test2() {
        try (IndexSharedStringTable sst = new IndexSharedStringTable()) {
            int length = 10000;
            String[] buf = new String[length];
            for (int i = 0; i < length; i++)
                buf[i] = getRandomString();

            for (String s : buf) {
                sst.push(s);
            }

            for (int i = 0; i < length; i++) {
                assert sst.get(i).equals(buf[i]);
            }

            int fromIndex = 0, size = length;
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

    @Test public void test3() throws IOException {
        IndexSharedStringTable sst = new IndexSharedStringTable();
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
        sst.commit();
    }
//
//    @Test public void test4() throws IOException {
//        Path path = Paths.get("/var/folders/rh/334bb3pn78s95dsn_tgvgyyw0000gn/T/+579019283137212671.sst.idx");
//        IndexSharedStringTable sst = new IndexSharedStringTable(path);
//        long start = System.currentTimeMillis();
//        int length = 512, n = 0;
//        String[] array = new String[length];
//        for (int i = 0; i < 10_000_000; ) {
//            int size = sst.batch(i, array);
//            i += size;
//            n += size;
//        }
//        System.out.println(n);
//        System.out.println(System.currentTimeMillis() - start);
//    }
}

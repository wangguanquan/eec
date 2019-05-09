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

import static cn.ttzero.excel.Print.println;
import static cn.ttzero.excel.Print.print;

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
}

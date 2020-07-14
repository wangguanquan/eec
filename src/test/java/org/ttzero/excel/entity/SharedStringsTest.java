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

package org.ttzero.excel.entity;

import org.junit.Test;

import java.io.IOException;


/**
 * @author guanquan.wang at 2019-05-07 17:41
 */
public class SharedStringsTest {
    @Test public void testGet() {
        try (SharedStrings sst = new SharedStrings()) {
            int index = sst.get("abc");
            assert index == 0;

            index = sst.get("guanquan.wang");
            assert index == 1;

            index = sst.get("abc");
            assert index == 0;

            index = sst.get("guanquan.wang");
            assert index == 1;

            index = sst.get("guanquan.wang");
            assert index == 1;

            index = sst.get("test");
            assert index == 2;
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    @Test public void testGetChar() {
        try (SharedStrings sst = new SharedStrings()) {
            for (int i = 0; i <= 0x7F; i++) {
                sst.get((char) i);
            }

            for (int i = 0; i < 0x7FFFFFFF; i++) {
                char c = (char) (i & 0x7F);
                int index = sst.get(c);
                assert index == c;
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    @Test public void testIterator() throws IOException {
//        Path path = Paths.get("/var/folders/rh/334bb3pn78s95dsn_tgvgyyw0000gn/T/+8037161714290441202.sst");
//        SharedStringTable sst = new SharedStringTable(path);
//
//        int i = 0;
//        for (String s : sst) {
//            print(i++);
//            print(' ');
//            println(s);
//        }
    }
}

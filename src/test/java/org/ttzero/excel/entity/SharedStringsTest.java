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
import java.util.HashMap;
import java.util.Map;

import static org.ttzero.excel.entity.WorkbookTest.author;
import static org.ttzero.excel.entity.WorkbookTest.getRandomString;


/**
 * @author guanquan.wang at 2019-05-07 17:41
 */
public class SharedStringsTest {
    @Test public void testGet() {
        try (SharedStrings sst = new SharedStrings()) {
            int index = sst.get("abc");
            assert index == 0;

            index = sst.get(author);
            assert index == 1;

            index = sst.get("abc");
            assert index == 0;

            index = sst.get(author);
            assert index == 1;

            index = sst.get(author);
            assert index == 1;

            index = sst.get("test");
            assert index == 2;
        } catch (IOException e) {
            e.printStackTrace();
            assert false;
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
            assert false;
        }
    }

    @Test public void testFind() throws IOException {
        try (SharedStrings sst = new SharedStrings()) {
            int size = 10_000;
            Map<String, Integer> indexMap = new HashMap<>(size);
            String v;
            for (int i = 0; i < size; i++) {
                v = getRandomString();

                indexMap.put(v, sst.get(v));
            }

            for (Map.Entry<String, Integer> entry : indexMap.entrySet()) {
                assert entry.getValue().equals(sst.get(entry.getKey()));
            }
        }
    }
}

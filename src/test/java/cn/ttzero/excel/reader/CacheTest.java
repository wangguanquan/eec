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

import java.util.Iterator;

import static cn.ttzero.excel.Print.println;

/**
 * Create by guanquan.wang at 2019-05-07 15:17
 */
public class CacheTest {
    @Test public void testPut1() {
        Cache<Integer, String> hot = FixSizeLRUCache.create();
        hot.put(1, "a");
        hot.put(2, "b");
        hot.put(3, "c");
        hot.put(4, "d");

        assert hot.size() == 4;
        assert hot.get(2).equals("b");

        hot.forEach(e -> println(e.getKey() + ": " + e.getValue()));

        println(hot.get(3));
    }

    @Test public void testPut2() {
        Cache<String, Integer> hot = FixSizeLRUCache.create();
        hot.put("a", 1);
        hot.put("b", 2);
        hot.put("a", 8);
        hot.put("c", 3);
        hot.put("d", 4);
        hot.put("e", 5);

        assert hot.size() == 5;
        assert hot.get("c") == 3;

        hot.forEach(e -> println(e.getKey() + ": " + e.getValue()));
    }

    @Test public void testIterator() {
        Cache<String, Integer> hot = FixSizeLRUCache.create();
        hot.put("a", 1);
        hot.put("b", 2);
        hot.put("c", 3);
        hot.put("d", 4);
        hot.put("e", 5);

        for (Iterator<Cache.Entry<String, Integer>> ite = hot.iterator(); ite.hasNext();) {
            Cache.Entry<String, Integer> e = ite.next();
            println(e.getKey() + ": " + e.getValue());
        }
    }

    @Test public void testRemove() {
        FixSizeLRUCache<Integer, String> hot = FixSizeLRUCache.create();
        hot.put(1, "a");
        hot.put(2, "b");
        hot.put(3, "c");
        hot.put(4, "d");

        assert hot.size() == 4;
        hot.forEach(e -> println(e.getKey() + ": " + e.getValue()));

        hot.remove();

        assert hot.size() == 3;

        hot.remove();
        hot.forEach(e -> println(e.getKey() + ": " + e.getValue()));
        hot.remove();
        hot.forEach(e -> println(e.getKey() + ": " + e.getValue()));
        hot.remove();
        hot.forEach(e -> println(e.getKey() + ": " + e.getValue()));
        hot.remove();
        hot.remove();
        hot.remove();
        hot.remove();
        hot.forEach(e -> println(e.getKey() + ": " + e.getValue()));
    }

    @Test public void testClear() {
        Cache<Integer, String> hot = FixSizeLRUCache.create();
        hot.put(1, "a");
        hot.put(2, "b");
        hot.put(3, "c");
        hot.put(4, "d");

        assert hot.size() == 4;
        hot.clear();

        assert hot.size() == 0;
        hot.forEach(e -> println(e.getKey() + ": " + e.getValue()));
    }
}

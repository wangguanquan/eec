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

package org.ttzero.excel.reader;

import org.junit.Test;

import static org.junit.Assert.assertNotEquals;
import static org.junit.Assert.assertNull;

/**
 * @author guanquan.wang at 2019-05-07 15:17
 */
public class CacheTest {
    @Test public void testPut1() {
        Cache<Integer, String> hot = FixSizeLRUCache.create();
        hot.put(1, "a");
        hot.put(2, "b");
        hot.put(3, "c");
        hot.put(4, "d");
        assert "4:d=>3:c=>2:b=>1:a".equals(hot.toString());

        assert hot.size() == 4;
        assert hot.get(2).equals("b");
        assert "2:b=>4:d=>3:c=>1:a".equals(hot.toString());
        assert hot.get(4).equals("d");
        assert "4:d=>2:b=>3:c=>1:a".equals(hot.toString());
        System.out.println();
    }

    @Test public void testPut2() {
        Cache<String, Integer> hot = FixSizeLRUCache.create();
        hot.put("a", 1);
        hot.put("b", 2);
        hot.put("a", 8);
        hot.put("c", 3);
        hot.put("d", 4);
        hot.put("e", 5);
        assert "e:5=>d:4=>c:3=>a:8=>b:2".equals(hot.toString());

        assert hot.size() == 5;
        assert hot.get("c") == 3;
        assert "c:3=>e:5=>d:4=>a:8=>b:2".equals(hot.toString());
    }

    @Test public void testIterator() {
        Cache<String, Integer> hot = FixSizeLRUCache.create();
        hot.put("a", 1);
        hot.put("b", 2);
        hot.put("c", 3);
        hot.put("d", 4);
        hot.put("e", 5);

        String[] expected = {"e:5", "d:4", "c:3", "b:2", "a:1"};
        int i = 0;
        for (Cache.Entry<String, Integer> e : hot) {
            assert expected[i++].equals(e.toString());
        }
    }

    @Test public void testRemoveTail() {
        FixSizeLRUCache<Integer, String> hot = FixSizeLRUCache.create();
        hot.put(1, "a");
        hot.put(2, "b");
        hot.put(3, "c");
        hot.put(4, "d");

        assert hot.size() == 4;
        assert hot.get(1).equals("a");
        assert "1:a=>4:d=>3:c=>2:b".equals(hot.toString());
        assertNotEquals(hot.get(2), "B");
        assert "2:b=>1:a=>4:d=>3:c".equals(hot.toString());

        hot.remove();

        assert "2:b=>1:a=>4:d".equals(hot.toString());
        assert hot.size() == 3;
        assert hot.get(1).equals("a");
        assert "1:a=>2:b=>4:d".equals(hot.toString());

        hot.remove();
        assert "1:a=>2:b".equals(hot.toString());
        assert hot.size() == 2;
        assert hot.get(2).equals("b");
        assert "2:b=>1:a".equals(hot.toString());
        hot.remove();
        assert "2:b".equals(hot.toString());
        assert hot.size() == 1;
        assertNull(hot.get(3));
        hot.remove();
        assert "".equals(hot.toString());
        assert hot.size() == 0;
        assertNull(hot.get(4));
        hot.remove();
        hot.remove();
        hot.remove();
        hot.remove();
        assert hot.size() == 0;
    }

    @Test public void testRemove() {
        FixSizeLRUCache<Integer, String> hot = FixSizeLRUCache.create();
        hot.put(1, "a");
        hot.put(2, "b");
        hot.put(3, "c");
        hot.put(4, "d");

        assert "b".equals(hot.remove(2));
        assert "4:d=>3:c=>1:a".equals(hot.toString());

        assert "a".equals(hot.remove(1));
        assert "4:d=>3:c".equals(hot.toString());

        assert "d".equals(hot.remove(4));
        assert "3:c".equals(hot.toString());

        assertNull(hot.remove(4));
        assert "3:c".equals(hot.toString());

        assert "c".equals(hot.remove(3));
        assert hot.size() == 0;
        assert "".equals(hot.toString());
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
    }

    @Test public void testRemoveAndAdd() {
        Cache<String, Integer> cache = FixSizeLRUCache.create();
        cache.put("a", 1);
        cache.put("b", 2);

        assert cache.size() == 2;
        assert cache.get("a") == 1;
        assert cache.get("b") == 2;

        cache.remove("a");
        assert cache.size() == 1;
        cache.put("a", 5);
        assert cache.size() == 2;
        assert cache.get("a") == 5;
    }
}

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

package cn.ttzero.excel.entity;

import cn.ttzero.excel.Print;
import org.junit.Test;

import java.io.IOException;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;
import java.util.Random;

/**
 * Create by guanquan.wang at 2019-04-26 17:40
 */
public class WorkbookTest {
    /**
     * The default output path
     */
    private Path defaultTestPath = Paths.get("target/excel/");
    private Random random = new Random();

    @Test public void testWrite() throws IOException {
        new Workbook("test", "guanquan.wang")
            .watch(Print::println)
            .addSheet(createTestData())
            .writeTo(defaultTestPath);
    }

    private List<Item> createTestData() {
        int n = random.nextInt(100) + 1;
        List<Item> list = new ArrayList<>(n);
        for (int i = 0; i < n; i++) {
            list.add(new Item(i, getRandom()));
        }
        return list;
    }

    private char[] charArray = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890".toCharArray();
    private char[][] cache = {new char[6], new char[7], new char[8], new char[9], new char[10]};
    public String getRandom() {
        int n = random.nextInt(5), size = charArray.length;
        char[] cs = cache[n];
        for (int i = 0; i < cs.length; i++) {
            cs[i] = charArray[random.nextInt(size)];
        }
        return new String(cs);
    }

    public static class Item {
        private int id;
        private String name;

        public Item(int id, String name) {
            this.id = id;
            this.name = name;
        }
    }
}

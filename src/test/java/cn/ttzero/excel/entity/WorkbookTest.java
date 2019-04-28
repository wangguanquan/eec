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

import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Random;

/**
 * Create by guanquan.wang at 2019-04-26 17:40
 */
public class WorkbookTest {
    /**
     * The default output path
     */
    Path defaultTestPath = Paths.get("target/excel/");
    static Random random = new Random();

    static char[] charArray = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890".toCharArray();
    private static char[][] cache = {new char[6], new char[7], new char[8], new char[9], new char[10]};
    static String getRandomString() {
        int n = random.nextInt(5), size = charArray.length;
        char[] cs = cache[n];
        for (int i = 0; i < cs.length; i++) {
            cs[i] = charArray[random.nextInt(size)];
        }
        return new String(cs);
    }

}

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

package org.ttzero.excel.entity;

import org.ttzero.excel.util.FileUtil;

import java.io.IOException;
import java.nio.file.Files;
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
    static Path defaultTestPath = Paths.get("target/excel/");
    String author = "guanquan.wang";
    static Random random = new Random();

    static char[] charArray = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890".toCharArray();
    private static char[] cache = new char[32];
    public static String getRandomString() {
        int n = random.nextInt(cache.length) + 1, size = charArray.length;
        for (int i = 0; i < n; i++) {
            cache[i] = charArray[random.nextInt(size)];
        }
        return new String(cache, 0, n);
    }

    public static Path getOutputTestPath() throws IOException {
        if (!Files.exists(defaultTestPath)) {
            FileUtil.mkdir(defaultTestPath);
        }
        return defaultTestPath;
    }

}

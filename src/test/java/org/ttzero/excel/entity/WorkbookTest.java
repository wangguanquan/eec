/*
 * Copyright (c) 2017-2019, guanquan.wang@hotmail.com All Rights Reserved.
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
import java.util.zip.CRC32;

import static org.ttzero.excel.util.FileUtil.exists;

/**
 * @author guanquan.wang at 2019-04-26 17:40
 */
public class WorkbookTest {
    /**
     * The default output path
     */
    public static Path defaultTestPath = Paths.get("target/excel/");
    public static String author = "eec";
    public static Random random = new Random();

    public static char[] charArray = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890❤不逢北国之秋，已将近十余年了。在南方每年到了秋天，总要想起陶然亭（1）的芦花，钓鱼台（2）的柳影，西山（3）的虫唱，玉泉（4）的夜月，潭柘寺（5）的钟声。在北平即使不出门去吧，就是在皇城人海之中，租人家一椽（6）破屋来住着，早晨起来，泡一碗浓茶，向院子一坐，你也能看得到很高很高的碧绿的天色，听得到青天下驯鸽的飞声。从槐树叶底，朝东细数着一丝一丝漏下来的日光，或在破壁腰中，静对着像喇叭似的牵牛花（朝荣）的蓝朵，自然而然地也能够感觉到十分的秋意。说到了牵牛花，我以为以蓝色或白色者为佳，紫黑色次之，淡红色最下。最好，还要在牵牛花底，叫长着几根疏疏落落的尖细且长的秋草，使作陪衬。".toCharArray();
    private static final char[] cache = new char[32];

    public static String getRandomAssicString() {
        return getRandomAssicString(20);
    }

    public static String getRandomAssicString(int maxlen) {
        int n = random.nextInt(maxlen) + 1, size = 62;
        for (int i = 0; i < n; i++) {
            cache[i] = charArray[random.nextInt(size)];
        }
        return new String(cache, 0, n);
    }

    public static String getRandomString(int maxLen) {
        int n = random.nextInt(maxLen) + 1, size = charArray.length;
        for (int i = 0; i < n; i++) {
            cache[i] = charArray[random.nextInt(size)];
        }
        return new String(cache, 0, n);
    }

    public static String getRandomString() {
        return getRandomString(cache.length);
    }

    public static Path getOutputTestPath() throws IOException {
        if (!exists(defaultTestPath)) {
            FileUtil.mkdir(defaultTestPath);
        }
        return defaultTestPath;
    }

    public static long crc32(Path path) {
        if (!Files.exists(path)) return 0L;
        CRC32 crc32 = new CRC32();
        try {
            crc32.update(Files.readAllBytes(path));
        } catch (IOException e) {
            return 0L;
        }
        return crc32.getValue();
    }

    public static long crc32(byte[] bytes) {
        CRC32 crc32 = new CRC32();
        crc32.update(bytes);
        return crc32.getValue();
    }
}

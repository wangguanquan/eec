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
import java.math.BigDecimal;
import java.sql.Time;
import java.sql.Timestamp;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.util.*;

/**
 * Create by guanquan.wang at 2019-04-28 19:16
 */
public class ListMapSheetTest extends WorkbookTest {

    @Test public void testWrite() throws IOException {
        new Workbook("test map", "guanquan.wang")
            .watch(Print::println)
            .addSheet(createTestData())
            .writeTo(defaultTestPath);
    }

    @Test public void testAllType() throws IOException {
        new Workbook("test map", "guanquan.wang")
                .watch(Print::println)
                .addSheet(createAllTypeData())
                .writeTo(defaultTestPath);
    }

    private List<Map<String, ?>> createTestData() {
        int size = random.nextInt(100) + 1;
        List<Map<String, ?>> list = new ArrayList<>(size);
        for (int i = 0; i < size; i++) {
            Map<String, Object> map = new HashMap<>();
            map.put("id", random.nextInt());
            map.put("name", getRandomString());
            list.add(map);
        }
        return list;
    }

    private List<Map<String, ?>> createAllTypeData() {
        int size = random.nextInt(100) + 1;
        List<Map<String, ?>> list = new ArrayList<>(size);
        for (int i = 0; i < size; i++) {
            Map<String, Object> map = new HashMap<>();
            map.put("cv", charArray[random.nextInt(charArray.length)]);
            map.put("sv", (short) (random.nextInt() & 0xFFFF));
            map.put("nv", random.nextInt());
            map.put("lv", random.nextLong());
            map.put("fv", random.nextFloat());
            map.put("dv", random.nextDouble());
            map.put("s", getRandomString());
            map.put("mv", BigDecimal.valueOf(random.nextDouble()));
            map.put("av", new Date());
            map.put("iv", new Timestamp(System.currentTimeMillis() - random.nextInt(9999999)));
            map.put("tv", new Time(random.nextLong()));
            map.put("ldv", LocalDate.now());
            map.put("ldtv", LocalDateTime.now());
            map.put("ltv", LocalTime.now());
            list.add(map);
        }
        return list;
    }
}

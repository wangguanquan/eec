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

import org.junit.Test;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.util.HashMap;
import java.util.Map;

import static cn.ttzero.excel.reader.ExcelReaderTest.testResourceRoot;

/**
 * Create by guanquan.wang at 2019-05-05 10:53
 */
public class TemplateTest extends WorkbookTest {

    @Test public void testTemplate() {
        try (InputStream fis = Files.newInputStream(testResourceRoot().resolve("template.xlsx"))) {
            // Map data
            Map<String, Object> map = new HashMap<>();
            map.put("name", "guanquan.wang");
            map.put("score", 90);
            map.put("date", "2019-05-05");
            map.put("desc", "暑假");

            // java bean
//            BindEntity entity = new BindEntity();
//            entity.score = 67;
//            entity.name = "张三";
//            entity.date = new Date(System.currentTimeMillis());

            new Workbook("模板导出", author)
                .withTemplate(fis, map)
                .writeTo(defaultTestPath);
        } catch (IOException | ExcelWriteException e) {
            e.printStackTrace();
        }
    }
}

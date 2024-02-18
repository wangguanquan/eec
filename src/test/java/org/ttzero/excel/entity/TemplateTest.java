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
import org.ttzero.excel.reader.ExcelReader;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.time.LocalDate;
import java.util.HashMap;
import java.util.Map;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;
import static org.ttzero.excel.reader.ExcelReaderTest.testResourceRoot;

/**
 * @author guanquan.wang at 2019-05-05 10:53
 */
public class TemplateTest extends WorkbookTest {

    @Test public void testTemplate() throws IOException {
        try (InputStream fis = Files.newInputStream(testResourceRoot().resolve("template.xlsx"))) {
            // Map data
            Map<String, Object> map = new HashMap<>();
            map.put("name", author);
            map.put("score", random.nextInt(90) + 10);
            map.put("date", LocalDate.now().toString());
            map.put("desc", "暑假");

            new Workbook()
                .withTemplate(fis, map)
                .writeTo(defaultTestPath.resolve("模板导出.xlsx"));

            try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("模板导出.xlsx"))) {
                reader.sheet(0).rows().forEach(row -> {
                    switch (row.getRowNum()) {
                        case 1:
                            assertEquals("通知书", row.getString(0).trim());
                            break;
                        case 3:
                            assertEquals((map.get("name") + " 同学，在本次期末考试的成绩是 " + map.get("score")+ "，希望"), row.getString(1).trim());
                            break;
                        case 4:
                            assertEquals(("下学期继续努力，祝你有一个愉快的" + map.get("desc") + "。"), row.getString(0).trim());
                            break;
                        case 23:
                            assertEquals(map.get("date"), row.getString(0).trim());
                            break;
                        default:
                            assertTrue(row.isBlank());
                    }
                });
            }
        }
    }
}

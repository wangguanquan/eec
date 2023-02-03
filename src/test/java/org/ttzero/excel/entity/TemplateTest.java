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
import org.ttzero.excel.annotation.ExcelColumn;
import org.ttzero.excel.annotation.FreezePanes;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import static org.ttzero.excel.reader.ExcelReaderTest.testResourceRoot;

/**
 * @author guanquan.wang at 2019-05-05 10:53
 */
public class TemplateTest extends WorkbookTest {

    @Test public void testTemplate() throws IOException {
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
        }
    }

    @Test public void testAutoSize() throws IOException {
        List<ServerReport> reports = new ArrayList<>(2);
        for (int i = 0; i < 2; i++) {
            ServerReport e = new ServerReport();
            e.index = i + 1;
            e.clientRequestNum = i;
            e.name = "测试测试测试测试测试测试测试测试测试测试测试测试测试测试测试测试";
            e.duration = 12345L;
            e.zipkinDt = "22-07-07";
            e.timestamp = new Date();
            reports.add(e);
        }
        new Workbook("服务数据")
            .setAutoSize(true)
            .addSheet(new ListSheet<>("服务报表1", reports)).writeTo(Paths.get("d://tmp/"));
    }


    @FreezePanes(topRow = 1)
    public static class ServerReport {

        @ExcelColumn(colIndex = 0, value = "序号")
        private Integer index;

        @ExcelColumn(colIndex = 2, value = "商户名")
        private String name;

        @ExcelColumn(colIndex = 1, value = "处理时间", hide = true, maxWidth = 254.88D)
        private String zipkinDt;

        @ExcelColumn(colIndex = 3, value = "时间", format = "yyyy-mm-dd hh:mm:ss")
        private Date timestamp;

        @ExcelColumn(colIndex = 4, value = "请求次数")
        private Integer clientRequestNum;

        @ExcelColumn(colIndex = 5, value = "持续时间")
        private Long duration;
    }
}

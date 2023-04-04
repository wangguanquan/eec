/*
 * Copyright (c) 2017-2023, guanquan.wang@yandex.com All Rights Reserved.
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
import org.ttzero.excel.manager.Const;

import java.io.IOException;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Date;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/**
 * @author guanquan.wang at 2023-02-04 22:15
 */
public class AutoSizeTest extends WorkbookTest {

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
            .addSheet(new ListSheet<>("服务报表1", reports))
            .writeTo(defaultTestPath);
    }

    @Test public void testAutoSize2() throws IOException {
        List<Map<String, ?>> reports = new ArrayList<>(2);
        for (int i = 0; i < 2; i++) {
            Map<String, Object> map = new LinkedHashMap<>();
            for (int j = 1; j <= Const.Limit.MAX_COLUMNS_ON_SHEET; j++) {
                map.put("COLUMN" + j, getRandomString());
            }
            reports.add(map);
        }
        new Workbook("服务数据")
            .setAutoSize(true)
            .addSheet(new ListMapSheet("服务报表1", reports)
            .putExtProp(Const.ExtendPropertyKey.FREEZE, Panes.row(1)))
            .writeTo(defaultTestPath);
    }

    @Test public void testAutoWidthAndFixedWidth() throws IOException {
        new Workbook("auto-width and fixed-width")
            .setAutoSize(true)
            .addSheet(new ListSheet<>(ListObjectSheetTest.Student.randomTestData()
                , new Column("学号", "id").fixedSize(16)
                , new Column("姓名", "name")
                , new Column("成绩", "score"))
            ).writeTo(defaultTestPath);
    }

    @Test public void testSpecifyColumnAutoWidth() throws IOException {
        new Workbook("specify column auto-width")
            .addSheet(new ListSheet<>(ListObjectSheetTest.Student.randomTestData()
                , new Column("学号", "id")
                , new Column("姓名", "name").autoSize()
                , new Column("成绩", "score")).fixedSize(10)
            ).writeTo(defaultTestPath);
    }

    @Test public void testFixedAndAutoWidth() throws IOException {
        new Workbook("fixed and fixed-width")
            .addSheet(new ListSheet<>(ListObjectSheetTest.Student.randomTestData()
                , new Column("学号", "id").fixedSize(10)
                , new Column("姓名", "name").autoSize()
                , new Column("成绩", "score"))
            ).writeTo(defaultTestPath);
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

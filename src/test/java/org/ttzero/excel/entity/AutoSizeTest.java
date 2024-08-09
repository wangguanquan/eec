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
import org.ttzero.excel.reader.ExcelReader;
import org.ttzero.excel.reader.Row;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.stream.Collectors;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;

/**
 * @author guanquan.wang at 2023-02-04 22:15
 */
public class AutoSizeTest extends WorkbookTest {

    @Test public void testAutoSize() throws IOException {
        List<ServerReport> expectList = new ArrayList<>(2);
        for (int i = 0; i < 2; i++) {
            ServerReport e = new ServerReport();
            e.index = i + 1;
            e.clientRequestNum = i;
            e.name = "测试测试测试测试测试测试测试测试测试测试测试测试测试测试测试测试";
            e.duration = 12345L;
            e.zipkinDt = "22-07-07";
            e.timestamp = new Date();
            expectList.add(e);
        }
        String fileName = "服务数据.xlsx";
        new Workbook()
            .setAutoSize(true)
            .addSheet(new ListSheet<>("服务报表1", expectList))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            List<ServerReport> list = reader.sheet(0).dataRows().map(row -> row.to(ServerReport.class)).collect(Collectors.toList());
            assertEquals(expectList.size(), list.size());
            for (int i = 0, len = expectList.size(); i < len; i++) {
                ServerReport expect = expectList.get(i), e = list.get(i);
                assertEquals(expect, e);
            }
        }
    }

    @Test public void testAutoSize2() throws IOException {
        List<Map<String, Object>> expectList = new ArrayList<>(2);
        for (int i = 0; i < 2; i++) {
            Map<String, Object> map = new LinkedHashMap<>();
            for (int j = 1; j <= Const.Limit.MAX_COLUMNS_ON_SHEET; j++) {
                map.put("COLUMN" + j, getRandomString());
            }
            expectList.add(map);
        }
        String fileName = "服务数据.xlsx";
        new Workbook()
            .setAutoSize(true)
            .addSheet(new ListMapSheet<>("服务报表1", expectList)
            .putExtProp(Const.ExtendPropertyKey.FREEZE, Panes.row(1)))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            List<Map<String, Object>> list = reader.sheet(0).dataRows().map(Row::toMap).collect(Collectors.toList());
            assertEquals(expectList.size(), list.size());
            for (int i = 0, len = expectList.size(); i < len; i++) {
                Map<String, Object> expect = expectList.get(i), e = list.get(i);
                assertEquals(expect, e);
            }
        }
    }

    @Test public void testAutoWidthAndFixedWidth() throws IOException {
        String fileName = "auto-width and fixed-width.xlsx";
        List<ListObjectSheetTest.Student> expectList = ListObjectSheetTest.Student.randomTestData();
        new Workbook()
            .setAutoSize(true)
            .addSheet(new ListSheet<>(expectList
                , new Column("学号", "id").fixedSize(16)
                , new Column("姓名", "name")
                , new Column("成绩", "score"))
            ).writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            List<ListObjectSheetTest.Student> list = reader.sheet(0).dataRows().map(row -> row.to(ListObjectSheetTest.Student.class)).collect(Collectors.toList());
            assertEquals(expectList.size(), list.size());
            for (int i = 0, len = expectList.size(); i < len; i++) {
                ListObjectSheetTest.Student expect = expectList.get(i), e = list.get(i);
                expect.setId(0);
                assertEquals(expect, e);
            }
        }
    }

    @Test public void testSpecifyColumnAutoWidth() throws IOException {
        String fileName = "specify column auto-width.xlsx";
        List<ListObjectSheetTest.Student> expectList = ListObjectSheetTest.Student.randomTestData();
        new Workbook()
            .addSheet(new ListSheet<>(expectList
                , new Column("学号", "id")
                , new Column("姓名", "name").autoSize()
                , new Column("成绩", "score")).fixedSize(10)
            ).writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            Iterator<org.ttzero.excel.reader.Row> iter = reader.sheet(0).dataIterator();
            for (ListObjectSheetTest.Student expect : expectList) {
                assertTrue(iter.hasNext());
                org.ttzero.excel.reader.Row row = iter.next();
                assertEquals((Integer) expect.getId(), row.getInt(0));
                assertEquals(expect.getName(), row.getString(1));
                assertEquals((Integer) expect.getScore(), row.getInt(2));
            }
        }
    }

    @Test public void testFixedAndAutoWidth() throws IOException {
        String fileName = "fixed and fixed-width.xlsx";
        List<ListObjectSheetTest.Student> expectList = ListObjectSheetTest.Student.randomTestData();
        new Workbook()
            .addSheet(new ListSheet<>(expectList
                , new Column("学号", "id")
                , new Column("姓名", "name").autoSize()
                , new Column("成绩", "score"))
            ).writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            Iterator<org.ttzero.excel.reader.Row> iter = reader.sheet(0).dataIterator();
            for (ListObjectSheetTest.Student expect : expectList) {
                assertTrue(iter.hasNext());
                org.ttzero.excel.reader.Row row = iter.next();
                assertEquals((Integer) expect.getId(), row.getInt(0));
                assertEquals(expect.getName(), row.getString(1));
                assertEquals((Integer) expect.getScore(), row.getInt(2));
            }
        }
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

        @Override
        public boolean equals(Object o) {
            if (this == o) return true;
            if (o == null || getClass() != o.getClass()) return false;
            ServerReport that = (ServerReport) o;
            return Objects.equals(index, that.index) &&
                Objects.equals(name, that.name) &&
                Objects.equals(zipkinDt, that.zipkinDt) &&
                timestamp.getTime() / 1000 == that.timestamp.getTime() / 1000 &&
                Objects.equals(clientRequestNum, that.clientRequestNum) &&
                Objects.equals(duration, that.duration);
        }

        @Override
        public int hashCode() {
            return Objects.hash(index, name, zipkinDt, timestamp, clientRequestNum, duration);
        }
    }
}

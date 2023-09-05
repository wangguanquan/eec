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


package org.ttzero.excel.reader;

import org.junit.Test;
import org.ttzero.excel.entity.ListMapSheet;
import org.ttzero.excel.entity.Workbook;

import java.io.IOException;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

import static org.ttzero.excel.entity.WorkbookTest.defaultTestPath;
import static org.ttzero.excel.reader.ExcelReaderTest.testResourceRoot;

/**
 * @author guanquan.wang at 2023-01-06 09:32
 */
public class ExcelReaderTest2 {
    @Test public void testIsBlank() throws IOException {
        try (ExcelReader reader = ExcelReader.read(testResourceRoot().resolve("#150.xlsx"))) {
            reader.sheet(0).rows().forEach(row -> {
                switch (row.getRowNum()) {
                    case 1:
                        assert !row.isEmpty();
                        assert !row.isBlank();
                        break;
                    case 2:
                        assert !row.isEmpty();
                        assert row.isBlank();
                        break;
                }
            });
        }
    }

    @Test public void test354() throws IOException {
        try (ExcelReader reader = ExcelReader.read(testResourceRoot().resolve("#354.xlsx"))) {
            List<Map<String, Object>> list = reader.sheet(0).dataRows().map(Row::toMap).collect(Collectors.toList());
            Map<String, Object> row1 = list.get(0);
            assert row1.get("通讯地址") != null;
            assert row1.get("紧急联系人姓名") != null;
            assert !"名字".equals(row1.get("通讯地址"));
            assert !"名字".equals(row1.get("紧急联系人姓名"));
        }
    }

    @Test public void testForceImport() throws IOException {
        Map<String, Object> data1 = new HashMap<>();
        data1.put("id", 1);
        data1.put("name", "abc");

        Map<String, Object> data2 = new HashMap<>();
        data2.put("id", 2);
        data2.put("name", "xyz");
        new Workbook()
            .addSheet(new ListMapSheet().setData(Arrays.asList(data1, data2)))
            .writeTo(defaultTestPath.resolve("Force Import.xlsx"));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("Force Import.xlsx"))) {
            List<U> list = reader.sheet(0).forceImport().dataRows().map(row -> row.to(U.class)).collect(Collectors.toList());
            assert list.size() == 2;
            assert "1: abc".equals(list.get(0).toString());
            assert "2: xyz".equals(list.get(1).toString());
        }
    }

    @Test public void testUpperCaseRead() throws IOException {
        Map<String, Object> data1 = new HashMap<>();
        data1.put("ID", 1);
        data1.put("NAME", "abc");

        Map<String, Object> data2 = new HashMap<>();
        data2.put("ID", 2);
        data2.put("NAME", "xyz");

        new Workbook()
            .addSheet(new ListMapSheet(Arrays.asList(data1, data2)))
            .writeTo(defaultTestPath.resolve("Upper case Reader test.xlsx"));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("Upper case Reader test.xlsx"))) {
            List<U> list = reader.sheet(0).forceImport().dataRows().map(row -> row.to(U.class)).collect(Collectors.toList());
            assert list.size() == 2;
            assert "0: null".equals(list.get(0).toString());
            assert "0: null".equals(list.get(1).toString());

            list = reader.sheet(0).reset().addHeaderColumnReadOption(HeaderRow.FORCE_IMPORT | HeaderRow.IGNORE_CASE)
                .dataRows().map(row -> row.to(U.class)).collect(Collectors.toList());
            assert list.size() == 2;
            assert "1: abc".equals(list.get(0).toString());
            assert "2: xyz".equals(list.get(1).toString());
        }
    }

    @Test public void testCamelCaseRead() throws IOException {
        Map<String, Object> data1 = new HashMap<>();
        data1.put("USER_ID", 1);
        data1.put("USER_NAME", "abc");

        Map<String, Object> data2 = new HashMap<>();
        data2.put("USER_ID", 2);
        data2.put("USER_NAME", "xyz");

        new Workbook()
            .addSheet(new ListMapSheet(Arrays.asList(data1, data2)))
            .writeTo(defaultTestPath.resolve("Underline case Reader test.xlsx"));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("Underline case Reader test.xlsx"))) {
            List<User> list = reader.sheet(0).forceImport().dataRows().map(row -> row.to(User.class)).collect(Collectors.toList());
            assert list.size() == 2;
            assert "0: null".equals(list.get(0).toString());
            assert "0: null".equals(list.get(1).toString());

            list = reader.sheet(0).reset().addHeaderColumnReadOption(HeaderRow.FORCE_IMPORT | HeaderRow.CAMEL_CASE)
                .dataRows().map(row -> row.to(User.class)).collect(Collectors.toList());
            assert list.size() == 2;
            assert "1: abc".equals(list.get(0).toString());
            assert "2: xyz".equals(list.get(1).toString());
        }
    }

    public static class U {
        int id;
        String name;

        public void setId(int id) {
            this.id = id;
        }

        @Override
        public String toString() {
            return id + ": " + name;
        }
    }

    public static class User {
        int userId;
        String userName;

        @Override
        public String toString() {
            return userId + ": " + userName;
        }
    }
}

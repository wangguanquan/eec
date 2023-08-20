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

import org.junit.FixMethodOrder;
import org.junit.runners.MethodSorters;
import org.junit.Test;
import org.ttzero.excel.reader.ExcelReader;

import java.io.IOException;

/**
 * @author guanquan.wang at 2019-04-29 21:36
 */
@FixMethodOrder(MethodSorters.NAME_ASCENDING)
public class EmptySheetTest extends WorkbookTest {
    @Test
    public void testEmpty() throws IOException {
        new Workbook("test empty", author)
            .addSheet(new EmptySheet())
            .writeTo(defaultTestPath);
    }

    @Test public void testEmptyWithHeader() throws IOException {
        new Workbook("test empty header", author)
            .addSheet(new EmptySheet("Empty"
                , new Column("ID", Integer.class)
                , new Column("NAME", String.class)
                , new Column("AGE", Integer.class)
            ))
            .writeTo(defaultTestPath);
    }

    @Test
    public void testEmptyReader() throws IOException {
        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("test empty.xlsx"))) {
            long count = reader.sheets().flatMap(org.ttzero.excel.reader.Sheet::rows).count();
            assert count == 0L;
        }
    }

    @Test
    public void testEmptyDataReader() throws IOException {
        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("test empty.xlsx"))) {
            long count = reader.sheets().flatMap(org.ttzero.excel.reader.Sheet::dataRows).count();
            assert count == 0L;
        }
    }

    @Test public void testEmptyWithHeaderReader() throws IOException {
        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("test empty header.xlsx"))) {
            long count = reader.sheets().flatMap(org.ttzero.excel.reader.Sheet::rows).count();
            assert count == 1L;
        }
    }

    @Test public void testEmptyWithHeaderDataReader() throws IOException {
        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("test empty header.xlsx"))) {
            long count = reader.sheets().flatMap(org.ttzero.excel.reader.Sheet::dataRows).count();
            assert count == 0L;
        }
    }
}

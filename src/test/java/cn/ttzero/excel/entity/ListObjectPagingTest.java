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
import cn.ttzero.excel.reader.ExcelReaderTest;
import org.junit.Test;

import java.io.IOException;

/**
 * Create by guanquan.wang at 2019-04-29 11:14
 */
public class ListObjectPagingTest extends WorkbookTest {

    @Test
    public void testPaging() throws IOException {
        new Workbook("test paging", author)
            .watch(Print::println)
            .addSheet(ListObjectSheetTest.Item.randomTestData(202))
            .setWorkbookWriter(new ReLimitXMLWorkbookWriter())
            .writeTo(defaultTestPath);
    }

    @Test
    public void testLessPaging() throws IOException {
        new Workbook("test less paging", author)
            .watch(Print::println)
            .addSheet(ListObjectSheetTest.Item.randomTestData(23))
            .setWorkbookWriter(new ReLimitXMLWorkbookWriter())
            .writeTo(defaultTestPath);
    }

    @Test public void testStringWaterMark() throws IOException {
        new Workbook("paging string water mark", author)
            .watch(Print::println)
            .setWaterMark(WaterMark.of("SECRET"))
            .addSheet(ListObjectSheetTest.Item.randomTestData())
            .setWorkbookWriter(new ReLimitXMLWorkbookWriter())
            .writeTo(defaultTestPath);
    }

    @Test public void testLocalPicWaterMark() throws IOException {
        new Workbook("paging local pic water mark", author)
            .watch(Print::println)
            .setWaterMark(WaterMark.of(ExcelReaderTest.testResourceRoot().resolve("mark.png")))
            .addSheet(ListObjectSheetTest.Item.randomTestData())
            .setWorkbookWriter(new ReLimitXMLWorkbookWriter())
            .writeTo(defaultTestPath);
    }

    @Test public void testStreamWaterMark() throws IOException {
        new Workbook("paging input stream water mark", author)
            .watch(Print::println)
            .setWaterMark(WaterMark.of(getClass().getClassLoader().getResourceAsStream("mark.png")))
            .addSheet(ListObjectSheetTest.Item.randomTestData())
            .setWorkbookWriter(new ReLimitXMLWorkbookWriter())
            .writeTo(defaultTestPath);
    }

    @Test public void testPagingCustomizeDataSource() throws IOException {
        new Workbook("paging customize datasource", author)
            .watch(Print::println)
            .setAutoSize(true)
            .addSheet(new CustomizeDataSourceSheet())
            .setWorkbookWriter(new ReLimitXMLWorkbookWriter())
            .writeTo(defaultTestPath);
    }
}

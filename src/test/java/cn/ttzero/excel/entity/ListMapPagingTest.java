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

/**
 * Create by guanquan.wang at 2019-04-29 11:14
 */
public class ListMapPagingTest extends WorkbookTest {

    @Test
    public void testPaging() throws IOException {
        new Workbook("test map paging", author)
            .watch(Print::println)
            .addSheet(ListMapSheetTest.createTestData(301))
            .setWorkbookWriter(new ReLimitXMLWorkbookWriter())
            .writeTo(defaultTestPath);
    }

    @Test
    public void testLessPaging() throws IOException {
        new Workbook("test map less paging", author)
            .watch(Print::println)
            .addSheet(ListMapSheetTest.createTestData(29))
            .setWorkbookWriter(new ReLimitXMLWorkbookWriter())
            .writeTo(defaultTestPath);
    }

}

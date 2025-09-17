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
import org.ttzero.excel.reader.HeaderRow;
import org.ttzero.excel.reader.Row;
import org.ttzero.excel.reader.Sheet;

import java.io.IOException;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import static org.junit.Assert.assertEquals;

/**
 * @author guanquan.wang at 2019-04-29 11:14
 */
public class ListMapPagingTest extends WorkbookTest {

    @Test public void testPaging() throws IOException {
        String fileName = "test map paging.xlsx";
        List<Map<String, Object>> expectList = ListMapSheetTest.createTestData(301);
        Workbook workbook = new Workbook()
            .addSheet(new ListMapSheet<>(expectList))
            .setWorkbookWriter(new ReLimitXMLWorkbookWriter());
        workbook.writeTo(defaultTestPath.resolve(fileName));

        int count = expectList.size(), rowLimit = workbook.getSheet(0).getSheetWriter().getRowLimit();
        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            assertEquals(reader.getSheetCount(), (count % (rowLimit - 1) > 0 ? count / (rowLimit - 1) + 1 : count / (rowLimit - 1))); // Include header row

            for (int i = 0, len = reader.getSheetCount(), a = 0; i < len; i++) {
                Sheet sheet = reader.sheet(i).header(1);
                org.ttzero.excel.reader.HeaderRow header = (HeaderRow) sheet.getHeader();
                assertEquals("id", header.get(0));
                assertEquals("name", header.get(1));
                Iterator<Row> iter = sheet.iterator();
                while (iter.hasNext()) {
                    Map<String, ?> expect = expectList.get(a++), e = iter.next().toMap();
                    assertEquals(expect, e);
                }
            }
        }
    }

    @Test public void testLessPaging() throws IOException {
        String fileName = "test map less paging.xlsx";
        List<Map<String, Object>> expectList = ListMapSheetTest.createTestData(29);
        Workbook workbook = new Workbook()
            .addSheet(new ListMapSheet<>(expectList))
            .setWorkbookWriter(new ReLimitXMLWorkbookWriter());
        workbook.writeTo(defaultTestPath.resolve(fileName));

        int count = expectList.size(), rowLimit = workbook.getSheet(0).getSheetWriter().getRowLimit();
        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            assertEquals(reader.getSheetCount(), (count % (rowLimit - 1) > 0 ? count / (rowLimit - 1) + 1 : count / (rowLimit - 1))); // Include header row

            for (int i = 0, len = reader.getSheetCount(), a = 0; i < len; i++) {
                Sheet sheet = reader.sheet(i).header(1);
                org.ttzero.excel.reader.HeaderRow header = (HeaderRow) sheet.getHeader();
                assertEquals("id", header.get(0));
                assertEquals("name", header.get(1));
                Iterator<Row> iter = sheet.iterator();
                while (iter.hasNext()) {
                    Map<String, ?> expect = expectList.get(a++), e = iter.next().toMap();
                    assertEquals(expect, e);
                }
            }
        }
    }

}

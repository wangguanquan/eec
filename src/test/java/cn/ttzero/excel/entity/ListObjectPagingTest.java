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
import cn.ttzero.excel.entity.e7.XMLWorkbookWriter;
import cn.ttzero.excel.entity.e7.XMLWorksheetWriter;
import cn.ttzero.excel.manager.Const;
import org.junit.Test;

import java.io.IOException;

/**
 * Create by guanquan.wang at 2019-04-29 11:14
 */
public class ListObjectPagingTest extends WorkbookTest {

    @Test
    public void testPaging() throws IOException {
        Workbook workbook = new Workbook("test paging", author)
            .watch(Print::println)
            .addSheet(ListObjectSheetTest.Item.randomTestData(202));
        workbook.setWorkbookWriter(new MyXMLWorkbookWriter(workbook))
            .writeTo(defaultTestPath);
    }

    @Test
    public void testLessPaging() throws IOException {
        Workbook workbook = new Workbook("test paging", author)
            .watch(Print::println)
            .addSheet(ListObjectSheetTest.Item.randomTestData(23));
        workbook.setWorkbookWriter(new MyXMLWorkbookWriter(workbook))
            .writeTo(defaultTestPath);
    }


    static class MyXMLWorkbookWriter extends XMLWorkbookWriter {

        public MyXMLWorkbookWriter(Workbook workbook) {
            super(workbook);
        }

        @Override
        protected IWorksheetWriter getWorksheetWriter(Sheet sheet) throws IOException {
            return new MyXMLWorksheetWriter(getWorkbook(), sheet);
        }
    }

    static class MyXMLWorksheetWriter extends XMLWorksheetWriter {

        public MyXMLWorksheetWriter(Workbook workbook, Sheet sheet) throws IOException {
            super(workbook, sheet);
        }

        /**
         * The Worksheet row limit
         * @return the limit
         */
        @Override
        public int getRowLimit() {
            return 50;
        }
    }
}

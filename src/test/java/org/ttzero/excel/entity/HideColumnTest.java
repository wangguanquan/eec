/*
 * Copyright (c) 2017-2022, guanquan.wang@yandex.com All Rights Reserved.
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

import java.io.IOException;
import java.util.List;

/**
 * @author guanquan.wang at 2022-08-03 08:50
 */
public class HideColumnTest extends WorkbookTest {
    @Test public void testColumnHide() throws IOException {
        Column[] columns = {new Column("ID", "id").hide(), new Column("NAME", "name")};
        new Workbook("Hide Column").addSheet(new ListSheet<>(ListObjectSheetTest.Item.randomTestData(10), columns))
            .writeTo(defaultTestPath);
    }

    @Test public void testColumnAnnoHide() throws IOException {
        new Workbook("Hide Column Annotation").addSheet(new ListSheet<>(HideColumnItem.randomTestData(10)))
            .writeTo(defaultTestPath);
    }

    public static class HideColumnItem extends ListObjectSheetTest.Item {
        public HideColumnItem() {
        }

        public HideColumnItem(int id, String name) {
            super(id, name);
        }

        @ExcelColumn(hide = true)
        @Override
        public int getId() {
            return super.getId();
        }

        public static List<ListObjectSheetTest.Item> randomTestData(int n) {
            return randomTestData(n, () -> new HideColumnItem(random.nextInt(100), getRandomString()));
        }
    }
}

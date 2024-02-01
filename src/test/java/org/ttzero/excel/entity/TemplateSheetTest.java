/*
 * Copyright (c) 2017-2024, guanquan.wang@yandex.com All Rights Reserved.
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
import org.ttzero.excel.reader.ExcelReaderTest;

import java.io.File;
import java.io.IOException;

/**
 * @author guanquan.wang at 2024-01-25 09:57
 */
public class TemplateSheetTest extends WorkbookTest {
    @Test public void testSimpleTemplate() throws IOException {
        String fileName = "simple template.xlsx";
        Workbook workbook = new Workbook();
        File[] files = ExcelReaderTest.testResourceRoot().toFile().listFiles();
        if (files != null) {
            for (File file : files) {
                if (file.getName().endsWith(".xlsx")) {
                    workbook.addSheet(new TemplateSheet(file.getName(), file.toPath()));
                }
            }
        }
        workbook.writeTo(getOutputTestPath().resolve(fileName));
    }
}

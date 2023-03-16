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

import java.io.IOException;

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
}

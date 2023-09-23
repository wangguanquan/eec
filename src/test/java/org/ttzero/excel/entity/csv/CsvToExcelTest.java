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

package org.ttzero.excel.entity.csv;

import org.junit.Before;
import org.junit.Test;
import org.ttzero.excel.entity.CSVSheet;
import org.ttzero.excel.entity.Workbook;
import org.ttzero.excel.entity.WorkbookTest;
import org.ttzero.excel.reader.ExcelReader;
import org.ttzero.excel.util.CSVUtil;
import org.ttzero.excel.util.StringUtil;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Iterator;
import java.util.List;

import static org.ttzero.excel.util.FileUtil.exists;

/**
 * @author guanquan.wang at 2019-09-26 10:07
 */
public class CsvToExcelTest extends WorkbookTest {
    private Path path;
    @Before public void before() throws IOException {
        path = getOutputTestPath().resolve("1.csv");

        // Create a test file
        if (!exists(path)) {
            int column = random.nextInt(10) + 1, row = random.nextInt(100) + 1;

            // storage column type
            // 0 string
            // 1 char
            // 2 int
            // 3 float
            // 4 double
            int[] types = new int[column];
            for (int i = 0; i < column; i++) {
                types[i] = random.nextInt(5);
            }

            try (CSVUtil.Writer writer = CSVUtil.newWriter(path)) {
                for (int i = 0; i < row; i++) {
                    for (int c = 0; c < column; c++) {
                        switch (types[c]) {
                            case 0:
                                writer.write(getRandomString());
                                break;
                            case 1:
                                writer.write(charArray[random.nextInt(charArray.length)]);
                                break;
                            case 2:
                                writer.write(random.nextInt());
                                break;
                            case 3:
                                writer.write(random.nextFloat());
                                break;
                            case 4:
                                writer.write(random.nextDouble());
                                break;
                        }
                    }
                    // break row
                    writer.newLine();
                }
            }
        }
    }

    @Test public void testFromPath() throws IOException {
        String fileName = "csv path test.xlsx";
        new Workbook()
            .addSheet(new CSVSheet(path))
            .writeTo(getOutputTestPath().resolve(fileName));

        List<String[]> expectList = CSVUtil.read(path);
        try (ExcelReader reader = ExcelReader.read(getOutputTestPath().resolve(fileName))) {
            Iterator<org.ttzero.excel.reader.Row> iter = reader.sheet(0).iterator();
            for (String[] expect : expectList) {
                assert iter.hasNext();
                org.ttzero.excel.reader.Row row = iter.next();
                for (int i = 0; i < expect.length; i++) {
                    if (expect[i] != null) {
                        assert expect[i].equals(row.getString(i));
                    } else {
                        assert StringUtil.isEmpty(row.getString(i));
                    }
                }
            }
        }
    }

    @Test public void testFromInputStream() throws IOException {
        String fileName = "csv inputstream test.xlsx";
        new Workbook()
            .addSheet(new CSVSheet(Files.newInputStream(path)))
            .writeTo(getOutputTestPath().resolve(fileName));

        List<String[]> expectList = CSVUtil.read(path);
        try (ExcelReader reader = ExcelReader.read(getOutputTestPath().resolve(fileName))) {
            Iterator<org.ttzero.excel.reader.Row> iter = reader.sheet(0).iterator();
            for (String[] expect : expectList) {
                assert iter.hasNext();
                org.ttzero.excel.reader.Row row = iter.next();
                for (int i = 0; i < expect.length; i++) {
                    if (expect[i] != null) {
                        assert expect[i].equals(row.getString(i));
                    } else {
                        assert StringUtil.isEmpty(row.getString(i));
                    }
                }
            }
        }
    }

    @Test public void testFromReader() throws IOException {
        String fileName = "csv reader test.xlsx";
        new Workbook()
            .addSheet(new CSVSheet(Files.newBufferedReader(path)))
            .writeTo(getOutputTestPath().resolve(fileName));

        List<String[]> expectList = CSVUtil.read(path);
        try (ExcelReader reader = ExcelReader.read(getOutputTestPath().resolve(fileName))) {
            Iterator<org.ttzero.excel.reader.Row> iter = reader.sheet(0).iterator();
            for (String[] expect : expectList) {
                assert iter.hasNext();
                org.ttzero.excel.reader.Row row = iter.next();
                for (int i = 0; i < expect.length; i++) {
                    if (expect[i] != null) {
                        assert expect[i].equals(row.getString(i));
                    } else {
                        assert StringUtil.isEmpty(row.getString(i));
                    }
                }
            }
        }
    }

    @Test public void testHasHeaderFromPath() throws IOException {
        String fileName = "csv header path test.xlsx";
        new Workbook()
            .addSheet(new CSVSheet(path).setHasHeader(true))
            .writeTo(getOutputTestPath().resolve(fileName));

        List<String[]> expectList = CSVUtil.read(path);
        try (ExcelReader reader = ExcelReader.read(getOutputTestPath().resolve(fileName))) {
            Iterator<org.ttzero.excel.reader.Row> iter = reader.sheet(0).iterator();
            for (String[] expect : expectList) {
                assert iter.hasNext();
                org.ttzero.excel.reader.Row row = iter.next();
                for (int i = 0; i < expect.length; i++) {
                    if (expect[i] != null) {
                        assert expect[i].equals(row.getString(i));
                    } else {
                        assert StringUtil.isEmpty(row.getString(i));
                    }
                }
            }
        }
    }
}

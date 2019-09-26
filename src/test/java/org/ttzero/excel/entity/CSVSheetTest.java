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

package org.ttzero.excel.entity;

import org.junit.Before;
import org.junit.Test;
import org.ttzero.excel.Print;
import org.ttzero.excel.util.CSVUtil;
import org.ttzero.excel.util.CSVUtilTest;

import java.io.IOException;
import java.net.URL;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

import static org.ttzero.excel.util.FileUtil.isWindows;

/**
 * Create by guanquan.wang at 2019-09-26 10:07
 */
public class CSVSheetTest extends WorkbookTest {
    private Path path;
    @Before
    public void before() {
        URL url = CSVUtilTest.class.getClassLoader().getResource(".");
        if (url == null) {
            throw new RuntimeException("Load test resources error.");
        }
        path = isWindows() ? Paths.get(url.getFile().substring(1)) : Paths.get(url.getFile());
        path = path.resolve("1.csv");

        // Create a test file
        if (!Files.exists(path)) {
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
            } catch (IOException e) {
                e.printStackTrace();
                assert false;
            }
        }
    }

    @Test public void testFromPath() throws IOException {
        new Workbook("csv path test", author)
            .addSheet(new CSVSheet(path))
            .writeTo(getOutputTestPath());
    }

    @Test public void testFromInputStream() throws IOException {
        new Workbook("csv inputstream test", author)
            .watch(Print::println)
            .addSheet(new CSVSheet(Files.newInputStream(path)))
            .writeTo(getOutputTestPath());
    }

    @Test public void testFromReader() throws IOException {
        new Workbook("csv reader test", author)
            .addSheet(new CSVSheet(Files.newBufferedReader(path)))
            .writeTo(getOutputTestPath());
    }

    @Test public void testHasHeaderFromPath() throws IOException {
        new Workbook("csv path test", author)
            .addSheet(new CSVSheet(path).setHasHeader(true))
            .writeTo(getOutputTestPath());
    }
}

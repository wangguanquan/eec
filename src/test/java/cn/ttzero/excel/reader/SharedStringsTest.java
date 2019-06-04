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

package cn.ttzero.excel.reader;

import cn.ttzero.excel.entity.WorkbookTest;
import cn.ttzero.excel.util.ExtBufferedWriter;
import cn.ttzero.excel.util.FileUtil;
import org.junit.After;
import org.junit.Before;
import org.junit.Test;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Arrays;
import java.util.List;

import static org.junit.Assert.assertEquals;

/**
 * Create by guanquan.wang at 2019-05-09 21:16
 */
public class SharedStringsTest {

    private Path root;
    private Path path;

    @Before public void before() throws IOException {
        root = WorkbookTest.getOutputTestPath();
        if (!Files.exists(root)) {
            FileUtil.mkdir(root);
        }
    }

    @After public void close() throws IOException {
        if (path != null) Files.delete(path);
    }

    @Test public void testGeneral() throws IOException {
        List<String> list = Arrays.asList("abc", "中文");
        writeTestData(list);
        try (SharedStrings sst = new SharedStrings(path, 0, 0).load()) {
            checkTrue(sst, list);
        }
    }

    @Test public void testEscape() throws IOException {
        List<String> list = Arrays.asList("<row>", "\"abc\"", "&nbsp;");
        writeTestData(list);
        try (SharedStrings sst = new SharedStrings(path, 0, 0).load()) {
            checkTrue(sst, list);
        }
    }

    private void checkTrue(SharedStrings sst, List<String> list) {
        for (int i = 0, size = list.size(); i < size; i++) {
            assertEquals(list.get(i), sst.get(i));
        }
    }

    private void writeTestData(List<String> list) throws IOException {
        path = root.resolve(WorkbookTest.getRandomString() + ".xml");
        try (ExtBufferedWriter writer = new ExtBufferedWriter(Files.newBufferedWriter(path))) {
            writer.write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            writer.newLine();
            writer.write("<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"");
            writer.writeInt(list.size());
            writer.write("\" uniqueCount=\"");
            writer.writeInt(list.size());
            writer.write("\">");

            for (String v : list) {
                writer.write("<si><t>");
                writer.escapeWrite(v);
                writer.write("</t></si>");
            }

            // Final
            writer.write("</sst>");
        }
    }
}

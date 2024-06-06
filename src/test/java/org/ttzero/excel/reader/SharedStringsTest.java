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

package org.ttzero.excel.reader;

import org.junit.After;
import org.junit.Before;
import org.junit.Test;
import org.ttzero.excel.util.ExtBufferedWriter;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Arrays;
import java.util.List;

import static org.junit.Assert.assertEquals;
import static org.ttzero.excel.entity.WorkbookTest.getOutputTestPath;
import static org.ttzero.excel.entity.WorkbookTest.getRandomString;

/**
 * @author guanquan.wang at 2019-05-09 21:16
 */
public class SharedStringsTest {

    private Path root;
    private Path path;

    @Before public void before() throws IOException {
        root = getOutputTestPath();
    }

    @After public void close() throws IOException {
        if (path != null) Files.delete(path);
    }

    @Test public void testGeneral() throws IOException {
        List<String> list = Arrays.asList("abc", "中文");
        writeTestData(list);
        try (SharedStrings sst = new SharedStrings(Files.newInputStream(path), 0, 0).load()) {
            checkTrue(sst, list);
        }
    }

    @Test public void testEscape() throws IOException {
        List<String> list = Arrays.asList("<row>", "\"abc\"", "&nbsp;", "<tag>",
            "random&more", "one'one", "with\"signs\"", "random&more",
            "<this will be escaped \ud83d\ude01>", "nothing+!()happens", "An 😀awesome 😃string with a few 😉emojis!");
        writeTestData(list);
        try (SharedStrings sst = new SharedStrings(Files.newInputStream(path), 0, 0).load()) {
            checkTrue(sst, list);
        }
    }

    @Test public void testEscape2() {
        char[] chars = "&lt;tag&gt;,random&amp;more,with&quot;signs&quot;,random&amp;more,&abcd;352,&lt;this will be escaped &#x1f601;&gt;,An &#128512;awesome &#128515;string with a few &#x1f609;emojis!".toCharArray();
        String desc = SharedStrings.escape(chars, 0, chars.length);
        assertEquals(desc, "<tag>,random&more,with\"signs\",random&more,&abcd;352,<this will be escaped \uD83D\uDE01>,An 😀awesome 😃string with a few 😉emojis!");
    }

    private void checkTrue(SharedStrings sst, List<String> list) {
        for (int i = 0, size = list.size(); i < size; i++) {
            assertEquals(list.get(i), sst.get(i));
        }
    }

    private void writeTestData(List<String> list) throws IOException {
        path = root.resolve(getRandomString() + ".xml");
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

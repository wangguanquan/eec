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
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

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

    @Test public void test354() throws IOException {
        try (ExcelReader reader = ExcelReader.read(testResourceRoot().resolve("#354.xlsx"))) {
            List<Map<String, Object>> list = reader.sheet(0).dataRows().map(Row::toMap).collect(Collectors.toList());
            Map<String, Object> row1 = list.get(0);
            assert row1.get("通讯地址") != null;
            assert row1.get("紧急联系人姓名") != null;
            assert !"名字".equals(row1.get("通讯地址"));
            assert !"名字".equals(row1.get("紧急联系人姓名"));
        }
    }

    @Test public void testMerge() throws IOException {
        try (ExcelReader reader = ExcelReader.read(testResourceRoot().resolve("merge.xlsx"))) {
            MergeSheet sheet = reader.sheet(0).asMergeSheet();
            List<Dimension> list = sheet.getMergeCells();
            assert list.size() == 4;
            assert list.get(0).equals(Dimension.of("B2:C2"));
            assert list.get(1).equals(Dimension.of("E5:F8"));
            assert list.get(2).equals(Dimension.of("A13:A20"));
            assert list.get(3).equals(Dimension.of("B16:E17"));

            sheet = reader.sheet(1).asMergeSheet();
            list = sheet.getMergeCells();
            assert list.size() == 2;
            assert list.get(0).equals(Dimension.of("BM2:BQ11"));
            assert list.get(1).equals(Dimension.of("A1:B26"));

            sheet = reader.sheet(2).asMergeSheet();
            list = sheet.getMergeCells();
            assert list.size() == 2;
            assert list.get(0).equals(Dimension.of("A16428:D16437"));
            assert list.get(1).equals(Dimension.of("A1:K3"));

            sheet = reader.sheet(3).asMergeSheet();
            list = sheet.getMergeCells();
            assert list.size() == 1;
            assert list.get(0).equals(Dimension.of("A1:CF1434"));
        }
    }

    @Test public void testLargeMerge() throws IOException {
        try (ExcelReader reader = ExcelReader.read(testResourceRoot().resolve("largeMerged.xlsx"))) {
            MergeSheet sheet = reader.sheet(0).asMergeSheet();
            List<Dimension> list = sheet.getMergeCells();
            assert list.size() == 2608;
            assert list.get(0).equals(Dimension.of("C3:F3"));
            assert list.get(1).equals(Dimension.of("J2:J3"));
            assert list.get(2).equals(Dimension.of("B2:B3"));
            assert list.get(3).equals(Dimension.of("C5:F5"));

            assert list.get(98).equals(Dimension.of("C82:F82"));
            assert list.get(120).equals(Dimension.of("A104:A106"));
            assert list.get(210).equals(Dimension.of("C176:F176"));
            assert list.get(984).equals(Dimension.of("C821:F821"));

            assert list.get(1626).equals(Dimension.of("B1362:B1371"));
            assert list.get(1627).equals(Dimension.of("J1362:J1363"));
            assert list.get(2381).equals(Dimension.of("B2006:B2007"));
            assert list.get(2396).equals(Dimension.of("J2019:J2020"));

            assert list.get(2596).equals(Dimension.of("C2190:F2190"));
            assert list.get(2601).equals(Dimension.of("J2195:J2196"));
            assert list.get(2605).equals(Dimension.of("C2198:F2198"));
            assert list.get(2607).equals(Dimension.of("C2200:F2200"));
        }
    }
}

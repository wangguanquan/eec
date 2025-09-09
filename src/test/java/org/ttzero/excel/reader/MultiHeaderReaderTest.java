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

import org.junit.Test;

import java.io.IOException;
import java.math.RoundingMode;
import java.util.List;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertNull;
import static org.junit.Assert.assertTrue;
import static org.ttzero.excel.reader.ExcelReaderTest.testResourceRoot;

/**
 * @author guanquan.wang at 2019-04-26 17:42
 */
public class MultiHeaderReaderTest {


    @Test public void testMergeExcel() throws IOException {
        try (ExcelReader reader = ExcelReader.read(testResourceRoot().resolve("merge.xlsx"))) {
            List<Dimension> list = reader.sheet(0).asFullSheet().getMergeCells();
            assertTrue(list.contains(Dimension.of("B2:C2")));
            assertTrue(list.contains(Dimension.of("E5:F8")));
            assertTrue(list.contains(Dimension.of("A13:A20")));
            assertTrue(list.contains(Dimension.of("B16:E17")));

            list.addAll(reader.sheet(1).asFullSheet().getMergeCells());
            assertTrue(list.contains(Dimension.of("A1:B26")));
            assertTrue(list.contains(Dimension.of("BM2:BQ11")));

            list.addAll(reader.sheet(2).asFullSheet().getMergeCells());
            assertTrue(list.contains(Dimension.of("A1:K3")));
            assertTrue(list.contains(Dimension.of("A16428:D16437")));

            list.addAll(reader.sheet(3).asFullSheet().getMergeCells());
            assertTrue(list.contains(Dimension.of("A1:CF1434")));
        }
    }

    @Test public void testMergeExcel2() throws IOException {
        try (ExcelReader reader = ExcelReader.read(testResourceRoot().resolve("#150.xlsx"))) {
            List<Dimension> list = reader.sheet(0).asFullSheet().getMergeCells();
            assertTrue(list.contains(Dimension.of("A2:A31")));
            assertTrue(list.contains(Dimension.of("B8:B13")));
            assertTrue(list.contains(Dimension.of("A48:A54")));
            assertTrue(list.contains(Dimension.of("B52:B54")));
        }
    }

    @Test public void testLargeMerge() throws IOException {
        try (ExcelReader reader = ExcelReader.read(testResourceRoot().resolve("largeMerged.xlsx"))) {
            FullSheet fullSheet = reader.sheet(0).asFullSheet();
            Grid grid = fullSheet.getMergeGrid();
            assertEquals(grid.size(), 2608);
            assertTrue(grid.test(3, 1));
            assertTrue(grid.test(382, 1));
            assertTrue(grid.test(722, 2));
            assertTrue(grid.test(1374, 2));
            assertTrue(grid.test(2101, 10));
            assertTrue(grid.test(2201, 6));
            assertFalse(grid.test(2113, 5));
            long count = fullSheet.rows().count();

            Sheet sheet = fullSheet.asSheet();
            assertEquals(sheet.getClass(), XMLSheet.class);
            assertEquals(sheet.reset().rows().count(), count);
        }
    }

    @Test public void testFractureMerged() throws IOException {
        try (ExcelReader reader = ExcelReader.read(testResourceRoot().resolve("fracture merged.xlsx"))) {
            Sheet sheet = reader.sheet(0);
            Row header = sheet.header(1, 2).getHeader();
            String headerString = header.toString();
            assertEquals("姓名 | 二级机构名称 | 三级机构名称 | 四级机构名称 | 参与次数 | 日均参与率(%) | 日均得分 | 2021-07-01:得分 | 2021-07-01:考试时长 | 2021-07-02:得分 | 2021-07-02:考试时长 | 2021-07-03:得分 | 2021-07-03:考试时长 | 2021-07-04:得分 | 2021-07-04:考试时长 | 2021-07-05:得分 | 2021-07-05:考试时长 | 2021-07-06:得分 | 2021-07-06:考试时长 | 2021-07-07:得分 | 2021-07-07:考试时长 | 2021-07-08:得分 | 2021-07-08:考试时长 | 2021-07-09:得分 | 2021-07-09:考试时长 | 2021-07-10:得分 | 2021-07-10:考试时长", headerString);

            sheet.rows().forEach(row -> {
                if (row.getRowNum() == 3) {
                    assertEquals(row.getString("姓名"), "张三1");
                    assertEquals((int) row.getInt("参与次数"), 7);
                    assertEquals(row.getDecimal("日均得分").setScale(2, RoundingMode.HALF_UP).toString(), "41.43");
                    assertEquals((int) row.getInt("2021-07-01:得分"), 30);
                    assertEquals((int) row.getInt("2021-07-01:考试时长"), 19);
                    assertNull(row.getInt("2021-07-04:得分"));
                    assertEquals((int) row.getInt("2021-07-09:得分"), 70);
                    assertEquals((int) row.getInt("2021-07-09:考试时长"), 20);
                    assertNull(row.getInt("2021-07-10:得分"));
                    assertNull(row.getInt("2021-07-10:考试时长"));
                } else if (row.getRowNum() == 4) {
                    assertEquals(row.getString("姓名"), "张三2");
                    assertEquals((int) row.getInt("参与次数"), 0);
                    assertNull(row.getDecimal("日均得分"));
                    assertNull(row.getInt("2021-07-01:得分"));
                    assertNull(row.getInt("2021-07-01:考试时长"));
                    assertNull(row.getInt("2021-07-04:得分"));
                    assertNull(row.getInt("2021-07-09:得分"));
                    assertNull(row.getInt("2021-07-09:考试时长"));
                    assertNull(row.getInt("2021-07-10:得分"));
                    assertNull(row.getInt("2021-07-10:考试时长"));
                }
            });
        }
    }

}

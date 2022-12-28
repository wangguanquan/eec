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
import org.ttzero.excel.Print;
import org.ttzero.excel.manager.Const;

import java.io.IOException;
import java.math.BigDecimal;
import java.util.Objects;

import static org.ttzero.excel.reader.ExcelReaderTest.testResourceRoot;

/**
 * @author guanquan.wang at 2019-04-26 17:42
 */
public class MultiHeaderReaderTest {


    @Test public void testMergeExcel() throws IOException {
        try (ExcelReader reader = ExcelReader.read(testResourceRoot().resolve("merge.xlsx"))) {
            reader.parseFormula().sheets().flatMap(Sheet::rows).forEach(Print::println);
        }
    }

    @Test public void testMergeExcel2() throws IOException {
        try (ExcelReader reader = ExcelReader.read(testResourceRoot().resolve("#150.xlsx"))) {
            reader.sheets().flatMap(Sheet::rows).forEach(Print::println);
        }
    }

    @Test public void testLargeMerge() throws IOException {
        try (ExcelReader reader = ExcelReader.read(testResourceRoot().resolve("largeMerged.xlsx"))) {
            MergeSheet mergeSheet = reader.sheet(0).asMergeSheet();
            Grid grid = mergeSheet.getMergeGrid();
            assert grid.size() == 2608;
            assert grid.test(3, 1);
            assert grid.test(382, 1);
            assert grid.test(722, 2);
            assert grid.test(1374, 2);
            assert grid.test(2101, 10);
            assert grid.test(2201, 6);
            assert !grid.test(2113, 5);
            long count = mergeSheet.rows().count();

            Sheet sheet = mergeSheet.asCalcSheet();
            assert sheet.getClass() == XMLCalcSheet.class;
            assert sheet.reset().rows().count() == count;

            sheet = sheet.asSheet();
            assert sheet.getClass() == XMLSheet.class;
            assert sheet.reset().rows().count() == count;

            sheet = sheet.asMergeSheet();
            assert sheet.getClass() == XMLMergeSheet.class;
            assert sheet.reset().rows().count() == count;
        }
    }

    @Test public void testFractureMerged() {
        try (ExcelReader reader = ExcelReader.read(testResourceRoot().resolve("fracture merged.xlsx"))) {
            Sheet sheet = reader.sheet(0);
            Row header = sheet.header(1, 2).getHeader();
            String headerString = header.toString();
            assert "姓名 | 二级机构名称 | 三级机构名称 | 四级机构名称 | 参与次数 | 日均参与率(%) | 日均得分 | 2021-07-01:得分 | 2021-07-01:考试时长 | 2021-07-02:得分 | 2021-07-02:考试时长 | 2021-07-03:得分 | 2021-07-03:考试时长 | 2021-07-04:得分 | 2021-07-04:考试时长 | 2021-07-05:得分 | 2021-07-05:考试时长 | 2021-07-06:得分 | 2021-07-06:考试时长 | 2021-07-07:得分 | 2021-07-07:考试时长 | 2021-07-08:得分 | 2021-07-08:考试时长 | 2021-07-09:得分 | 2021-07-09:考试时长 | 2021-07-10:得分 | 2021-07-10:考试时长".equals(headerString.substring(0, headerString.indexOf(Const.lineSeparator)));

            sheet.rows().forEach(row -> {
                if (row.getRowNum() == 3) {
                    assert row.getString("姓名").equals("张三1");
                    assert row.getInt("参与次数").equals(7);
                    assert row.getDecimal("日均得分").setScale(2, BigDecimal.ROUND_HALF_DOWN).toString().equals("41.43");
                    assert row.getInt("2021-07-01:得分").equals(30);
                    assert row.getInt("2021-07-01:考试时长").equals(19);
                    assert Objects.isNull(row.getInt("2021-07-04:得分"));
                    assert row.getInt("2021-07-09:得分").equals(70);
                    assert row.getInt("2021-07-09:考试时长").equals(20);
                    assert Objects.isNull(row.getInt("2021-07-10:得分"));
                    assert Objects.isNull(row.getInt("2021-07-10:考试时长"));
                } else if (row.getRowNum() == 4) {
                    assert row.getString("姓名").equals("张三2");
                    assert row.getInt("参与次数").equals(0);
                    assert Objects.isNull(row.getDecimal("日均得分"));
                    assert Objects.isNull(row.getInt("2021-07-01:得分"));
                    assert Objects.isNull(row.getInt("2021-07-01:考试时长"));
                    assert Objects.isNull(row.getInt("2021-07-04:得分"));
                    assert Objects.isNull(row.getInt("2021-07-09:得分"));
                    assert Objects.isNull(row.getInt("2021-07-09:考试时长"));
                    assert Objects.isNull(row.getInt("2021-07-10:得分"));
                    assert Objects.isNull(row.getInt("2021-07-10:考试时长"));
                }
            });
        } catch (IOException ex) {
            ex.printStackTrace();
        }
    }

}

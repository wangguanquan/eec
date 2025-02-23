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


package org.ttzero.excel.entity;

import org.junit.Test;
import org.ttzero.excel.manager.Const;
import org.ttzero.excel.reader.Dimension;
import org.ttzero.excel.util.DateUtil;
import org.ttzero.excel.validation.DateValidation;
import org.ttzero.excel.validation.ListValidation;
import org.ttzero.excel.validation.TimeValidation;
import org.ttzero.excel.validation.Validation;
import org.ttzero.excel.validation.WholeValidation;

import java.io.IOException;
import java.sql.Time;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;


/**
 * @author guanquan.wang at 2022-08-17 20:05:42
 */
public class ValidationTest extends WorkbookTest {
    @Test public void testValidation() throws IOException {
        List<Validation> expectValidations = new ArrayList<>();
        // 下拉框选择“男”，“女”
        expectValidations.add(new ListValidation<>().in("男", "女").dimension(Dimension.of("A1")));
        // B1:E1 单元格只能输入大于1的数
        expectValidations.add(new WholeValidation().greaterThan(1).dimension(Dimension.of("B1:E1")));
        // 限制日期在2022年
        expectValidations.add(new DateValidation().between("2022-01-01", "2022-12-31").dimension(Dimension.of("A2")));
        // 限制时间小于下午6点（因为此时下班...）
        expectValidations.add(new TimeValidation().lessThan(DateUtil.toTimeValue(Time.valueOf("18:00:00"))).dimension(Dimension.of("B2")));
        // 引用
        expectValidations.add(new ListValidation<>().in(Dimension.of("A10:A12")).dimension(Dimension.of("D5")));
        final String fileName = "Validation Test.xlsx";
        new Workbook()
            .addSheet(new ListSheet<>(Arrays.asList("北京", "天津", "上海"))
                .setStartRowIndex(10, false)
                .putExtProp(Const.ExtendPropertyKey.DATA_VALIDATION, expectValidations))
            .writeTo(defaultTestPath.resolve(fileName));
    }

    @Test public void testValidationExtension() throws IOException {
        List<Validation> expectValidations = new ArrayList<>();
        // 引用
        expectValidations.add(new ListValidation<>().in("Sheet2", Dimension.of("A1:A3")).dimension(Dimension.of("D1:D5")));
        final String fileName = "Validation Extension Test.xlsx";
        new Workbook()
            .addSheet(new ListSheet<>().putExtProp(Const.ExtendPropertyKey.DATA_VALIDATION, expectValidations))
            .addSheet(new ListSheet<>(Arrays.asList("未知","男","女")))
            .writeTo(defaultTestPath.resolve(fileName));
    }
}

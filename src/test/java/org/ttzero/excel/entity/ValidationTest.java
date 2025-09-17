/*
 * Copyright (c) 2017-2019, guanquan.wang@hotmail.com All Rights Reserved.
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
        expectValidations.add(new ListValidation<>().in("未知", "男", "女").dimension(Dimension.of("A1")).setPrompt("未输入时默认\"未知\""));
        // B1:E1 单元格只能输入大于1的数
        expectValidations.add(new WholeValidation().greaterThan(1L).dimension(Dimension.of("B1:E1")).setPrompt("只能输入>1的数"));
        // 限制日期在2022年
        expectValidations.add(new DateValidation().between("2022-01-01", "2022-12-31").dimension(Dimension.of("A2")).setPrompt("只能输入\"2022-01-01\"~\"2022-12-31\""));
        // 限制时间小于下午6点（因为此时下班...）
        expectValidations.add(new TimeValidation().lessThan(DateUtil.toTimeValue(Time.valueOf("18:00:00"))).dimension(Dimension.of("B2")));
        // 引用
        expectValidations.add(new ListValidation<>().in(Dimension.of("A10:A12")).dimension(Dimension.of("D5")));

        expectValidations.add(ListValidation.in(Dimension.of("H2"), Arrays.asList("A省", "B省", "D省"))
            .addCascadeList(Dimension.of("I2:"), new LinkedHashMap<String, List<String>>(){{
                put("A省", Arrays.asList("A1市", "A2市", "A3市"));
                put("B省", Arrays.asList("B1市", "B2市", "B3市", "B4市"));
                put("D省", Arrays.asList("D1市", "D2市"));
            }})
            .addCascadeList(Dimension.of("J2:"), new LinkedHashMap<String, List<String>>(){{
                put("A1市", Arrays.asList("A1市-1", "A1市-2", "A1市-3"));
                put("A2市", Arrays.asList("A2市-1", "A2市-2", "A2市-3", "A2市-4", "A2市-5"));
                put("A3市", Arrays.asList("A3区-1", "A3区-2"));
                put("B1市", Arrays.asList("B1市-1", "B1市-2"));
                put("B2市", Arrays.asList("B2市-1", "B2市-2", "B2市-3"));
                put("B3市", Arrays.asList("B3市-1", "B3市-2", "B3市-3", "B3市-4"));
                put("B4市", Arrays.asList("B4市-1", "B4市-2"));
                put("D1市", Arrays.asList("D1市-1", "D1市-2"));
                put("D2市", Arrays.asList("D2市-1", "D2市-2", "D2市-3", "D2市-4"));
            }})
        );

        final String fileName = "Validation Test.xlsx";
        new Workbook()
            .addSheet(new ListSheet<>(Arrays.asList("北京", "天津", "上海"))
                .setStartCoordinate(10)
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

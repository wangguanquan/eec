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
import org.ttzero.excel.reader.CrossDimension;
import org.ttzero.excel.reader.Dimension;
import org.ttzero.excel.reader.ExcelReader;
import org.ttzero.excel.reader.FullSheet;
import org.ttzero.excel.reader.Row;
import org.ttzero.excel.util.DateUtil;
import org.ttzero.excel.validation.DateValidation;
import org.ttzero.excel.validation.DecimalValidation;
import org.ttzero.excel.validation.ListValidation;
import org.ttzero.excel.validation.TimeValidation;
import org.ttzero.excel.validation.Validation;
import org.ttzero.excel.validation.WholeValidation;

import java.io.IOException;
import java.math.BigDecimal;
import java.sql.Time;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertTrue;
import static org.ttzero.excel.reader.ExcelReaderTest.testResourceRoot;


/**
 * @author guanquan.wang at 2022-08-17 20:05:42
 */
public class ValidationTest extends WorkbookTest {
    @Test public void testValidation() throws IOException {
        List<Validation> expectValidations = new ArrayList<>();
        // 下拉框选择“男”，“女”
        expectValidations.add(new ListValidation<>().in("未知", "男", "女").dimension(Dimension.of("A1")).prompt("未输入时默认\"未知\""));
        // B1:E1 单元格只能输入大于1的数
        expectValidations.add(new WholeValidation().greaterThan(1L).dimension(Dimension.of("B1:E1")).prompt("只能输入>1的数"));
        // 限制日期在2022年
        expectValidations.add(new DateValidation().between("2022-01-01", "2022-12-31").dimension(Dimension.of("A2")).prompt("只能输入\"2022-01-01\"~\"2022-12-31\""));
        // 限制时间小于下午6点（因为此时下班...）
        expectValidations.add(new TimeValidation().lessThan(DateUtil.toTimeValue(Time.valueOf("18:00:00"))).dimension(Dimension.of("B2")));
        // 引用
        expectValidations.add(new ListValidation<>().referer(new CrossDimension(Dimension.of("A10:A12"))).dimension(Dimension.of("D5")));

        final String fileName = "Validation Test.xlsx";
        new Workbook()
            .addSheet(new ListSheet<>(Arrays.asList("北京", "天津", "上海"))
                .setStartCoordinate(10)
                .putExtProp(Const.ExtendPropertyKey.DATA_VALIDATION, expectValidations))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            FullSheet sheet = reader.sheet(0).asFullSheet();
            List<Validation> validations = sheet.getValidations();
            assertEquals(validations.size(), expectValidations.size());
        }
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

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            FullSheet sheet = reader.sheet(0).asFullSheet();
            List<Validation> validations = sheet.getValidations();
            assertEquals(validations.size(), expectValidations.size());
            Validation val = validations.get(0), expect = expectValidations.get(0);
            assertEquals(val.getType(), expect.getType());
            assertEquals(val.referer, expect.referer);
            assertEquals(val.sqrefList.size(), expect.sqrefList.size());
            for (Dimension sqref : val.sqrefList) {
                assertTrue(expect.sqrefList.contains(sqref));
            }
        }
    }

    @Test public void testTailElements() throws IOException {
        try (ExcelReader reader = ExcelReader.read(testResourceRoot().resolve("template2.xlsx"))) {
            FullSheet sheet = reader.sheet(2).asFullSheet();
            List<Dimension> mergeCells = sheet.getMergeCells();
            assertNotNull(mergeCells);
            assertEquals(mergeCells.size(), 3);
            assertTrue(mergeCells.contains(Dimension.of("E2:F11")));
            assertTrue(mergeCells.contains(Dimension.of("G2:N2")));
            assertTrue(mergeCells.contains(Dimension.of("E12:F12")));
            List<Validation> validations = sheet.getValidations();
            assertNotNull(validations);
            assertEquals(validations.size(), 15);
            Validation validation0 = validations.get(0), expectValidation0 = new ListValidation<>().in("自营德国海外仓","美国自营海外仓").dimension(Dimension.of("D14"));
            assertEquals(validation0.toString(), expectValidation0.toString());
            Validation validation1 = validations.get(1), expectValidation1 = new ListValidation<>().in(Dimension.of("F10:F12")).dimension(Dimension.of("F12:F13"));
            assertEquals(validation1.toString(), expectValidation1.toString());
            Validation validation9 = validations.get(9), expectValidation9 = new DecimalValidation().between(new BigDecimal("0.00001"), new BigDecimal("99999")).error("警告","只能填数字", Validation.ErrorStyle.warning).dimension(Dimension.of("L1"));
            assertEquals(validation9.toString(), expectValidation9.toString());
            Validation validation11 = validations.get(11), expectValidation11 = new WholeValidation().between(1L, 11111111L).dimension(Dimension.of("J1"));
            assertEquals(validation11.toString(), expectValidation11.toString());
        }
    }

    @Test public void testCascadeList() throws IOException {
        List<Validation> expectValidations = new ArrayList<>();
        expectValidations.add(ListValidation.in(Dimension.of("A2:"), Arrays.asList("A省", "B省", "D省"))
            .addCascadeList(Dimension.of("B2:"), new LinkedHashMap<String, List<String>>(){{
                put("A省", Arrays.asList("A1市", "A2市", "A3市"));
                put("B省", Arrays.asList("B1市", "B2市", "B3市", "B4市"));
                put("D省", Arrays.asList("D1市", "D2市"));
            }})
            .addCascadeList(Dimension.of("C2:"), new LinkedHashMap<String, List<String>>(){{
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
        final String fileName = "Validation CascadeList Test.xlsx";
        new Workbook()
            .addSheet(new SimpleSheet<>(Collections.singletonList(Arrays.asList("省份","市区","市镇"))).putExtProp(Const.ExtendPropertyKey.DATA_VALIDATION, expectValidations))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            FullSheet sheet = reader.sheet(0).asFullSheet();
            List<Validation> validations = sheet.getValidations();
            assertNotNull(validations);
            assertEquals(validations.size(), 2);
        }
    }

    @Test public void testListValidationOverSize() throws IOException {
        List<String> addressList = Arrays.asList("东山县","南靖县","龙海区","诏安县","龙文区","平和县","漳浦县","芗城区","云霄县","华安县","长泰县","南安市","石狮市","泉港区","金门县","丰泽区","安溪县","永春县","洛江区","德化县","晋江市","鲤城区","惠安县","将乐县","沙县区","永安市","梅列区","尤溪县","建宁县","宁化县","三元区","明溪县","泰宁县","大田县","清流县","古田县","寿宁县","福安市","周宁县","蕉城区","霞浦县","福鼎市","屏南县","柘荣县","同安区","翔安区","思明区","湖里区","集美区","海沧区","仙游县","秀屿区","荔城区","城厢区","涵江区","鼓楼区","平潭县","长乐区","闽清县","福清市","连江县","晋安区","罗源县","闽侯县","仓山区","永泰县","马尾区","台江区","邵武市","建阳区","建瓯市","武夷山市","延平区","顺昌县","光泽县","政和县","浦城县","松溪县","上杭县","武平县","漳平市","新罗区","连城县","永定区","长汀县");
        List<Validation> expectValidations = new ArrayList<>();
        expectValidations.add(new ListValidation<String>().in(addressList).dimension(Dimension.of("A1")).prompt("请选择地址"));
        final String fileName = "list oversize Test .xlsx";
        new Workbook()
            .addSheet(new ListSheet<>().putExtProp(Const.ExtendPropertyKey.DATA_VALIDATION, expectValidations))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            assertEquals(reader.getSheetCount(), 2);
            FullSheet sheet = reader.sheet(0).asFullSheet();
            List<Validation> validations = sheet.getValidations();
            assertEquals(validations.size(), expectValidations.size());

            org.ttzero.excel.reader.Sheet sheet2 = reader.sheet(1);
            // 判断是否为引用
            Validation expectVal = new ListValidation<>().dimension(Dimension.of("A1")).prompt("请选择地址").referer(new CrossDimension(sheet2.getName(), new Dimension(1, (short) 1, 1, (short) addressList.size())));
            assertEquals(validations.get(0).toString(), expectVal.toString());

            // 判断选项内容和顺序
            Iterator<Row> iter = sheet2.iterator();
            assertTrue(iter.hasNext());
            Row row0 = iter.next();
            assertTrue(addressList.size() >= row0.getLastColumnIndex());
            int i = 0;
            for (; i < row0.getLastColumnIndex(); i++) {
                assertEquals(addressList.get(i), row0.getString(i));
            }
            assertEquals(addressList.size(), i);
        }
    }
}

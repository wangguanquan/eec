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
            for (int i = 0; i < validations.size(); i++) {
                Validation expectVal = expectValidations.get(i), readVal = validations.get(i);
                assertEquals(expectVal.toString(), readVal.toString());
            }
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
                put("A1市", Arrays.asList("A1市_1", "A1市_2", "A1市_3"));
                put("A2市", Arrays.asList("A2市_1", "A2市_2", "A2市_3", "A2市_4", "A2市_5"));
                put("A3市", Arrays.asList("A3区_1", "A3区_2"));
                put("B1市", Arrays.asList("B1市_1", "B1市_2"));
                put("B2市", Arrays.asList("B2市_1", "B2市_2", "B2市_3"));
                put("B3市", Arrays.asList("B3市_1", "B3市_2", "B3市_3", "B3市_4"));
                put("B4市", Arrays.asList("B4市_1", "B4市_2"));
                put("D1市", Arrays.asList("D1市_1", "D1市_2"));
                put("D2市", Arrays.asList("D2市_1", "D2市_2", "D2市_3", "D2市_4"));
            }})
            .addCascadeList(Dimension.of("D2:"), new LinkedHashMap<String, List<String>>() {{
                put("A1市_1", Arrays.asList("A1市_1_1", "A1市_1_2"));
                put("A1市_2", Arrays.asList("A1市_2_1", "A1市_2_2"));
                put("A1市_3", Arrays.asList("A1市_3_1", "A1市_3_2"));
                put("A2市_1", Arrays.asList("A2市_1_1", "A2市_1_2"));
                put("A2市_2", Arrays.asList("A2市_2_1", "A2市_2_2"));
                put("A2市_3", Arrays.asList("A2市_3_1", "A2市_3_2"));
                put("A2市_4", Arrays.asList("A2市_4_1", "A2市_4_2"));
                put("A2市_5", Arrays.asList("A2市_5_1", "A2市_5_2"));
                put("A3区_1", Arrays.asList("A3区_1_1", "A3区_1_2"));
                put("A3区_2", Arrays.asList("A3区_2_1", "A3区_2_2"));
                put("B1市_1", Arrays.asList("B1市_1_1", "B1市_1_2"));
                put("B1市_2", Arrays.asList("B1市_2_1", "B1市_2_2"));
                put("B2市_1", Arrays.asList("B2市_1_1", "B2市_1_2"));
                put("B2市_2", Arrays.asList("B2市_2_1", "B2市_2_2"));
                put("B2市_3", Arrays.asList("B2市_3_1", "B2市_3_2"));
                put("B3市_1", Arrays.asList("B3市_1_1", "B3市_1_2"));
                put("B3市_2", Arrays.asList("B3市_2_1", "B3市_2_2"));
                put("B3市_3", Arrays.asList("B3市_3_1", "B3市_3_2"));
                put("B3市_4", Arrays.asList("B3市_4_1", "B3市_4_2"));
                put("B4市_1", Arrays.asList("B4市_1_1", "B4市_1_2"));
                put("B4市_2", Arrays.asList("B4市_2_1", "B4市_2_2"));
                put("D1市_1", Arrays.asList("D1市_1_1", "D1市_1_2"));
                put("D1市_2", Arrays.asList("D1市_2_1", "D1市_2_2"));
                put("D2市_1", Arrays.asList("D2市_1_1", "D2市_1_2"));
                put("D2市_2", Arrays.asList("D2市_2_1", "D2市_2_2"));
                put("D2市_3", Arrays.asList("D2市_3_1", "D2市_3_2"));
                put("D2市_4", Arrays.asList("D2市_4_1", "D2市_4_2"));
            }})
            .addCascadeList(Dimension.of("E2:"), new LinkedHashMap<String, List<String>>() {{
                put("A1市_1_1", Arrays.asList("A1市_1_1_1", "A1市_1_1_2", "其它"));
                put("A1市_1_2", Arrays.asList("A1市_1_2_1", "A1市_1_2_2", "其它"));
                put("A1市_2_1", Arrays.asList("A1市_2_1_1", "A1市_2_1_2", "其它"));
                put("A1市_2_2", Arrays.asList("A1市_2_2_1", "A1市_2_2_2", "其它"));
                put("A1市_3_1", Arrays.asList("A1市_3_1_1", "A1市_3_1_2", "其它"));
                put("A1市_3_2", Arrays.asList("A1市_3_2_1", "A1市_3_2_2", "其它"));
                put("A2市_1_1", Arrays.asList("A2市_1_1_1", "A2市_1_1_2", "其它"));
                put("A2市_1_2", Arrays.asList("A2市_1_2_1", "A2市_1_2_2", "其它"));
                put("A2市_2_1", Arrays.asList("A2市_2_1_1", "A2市_2_1_2", "其它"));
                put("A2市_2_2", Arrays.asList("A2市_2_2_1", "A2市_2_2_2", "其它"));
                put("A2市_3_1", Arrays.asList("A2市_3_1_1", "A2市_3_1_2", "其它"));
                put("A2市_3_2", Arrays.asList("A2市_3_2_1", "A2市_3_2_2", "其它"));
                put("A2市_4_1", Arrays.asList("A2市_4_1_1", "A2市_4_1_2", "其它"));
                put("A2市_4_2", Arrays.asList("A2市_4_2_1", "A2市_4_2_2", "其它"));
                put("A2市_5_1", Arrays.asList("A2市_5_1_1", "A2市_5_1_2", "其它"));
                put("A2市_5_2", Arrays.asList("A2市_5_2_1", "A2市_5_2_2", "其它"));
                put("A3区_1_1", Arrays.asList("A3区_1_1_1", "A3区_1_1_2", "其它"));
                put("A3区_1_2", Arrays.asList("A3区_1_2_1", "A3区_1_2_2", "其它"));
                put("A3区_2_1", Arrays.asList("A3区_2_1_1", "A3区_2_1_2", "其它"));
                put("A3区_2_2", Arrays.asList("A3区_2_2_1", "A3区_2_2_2", "其它"));
                put("B1市_1_1", Arrays.asList("B1市_1_1_1", "B1市_1_1_2", "其它"));
                put("B1市_1_2", Arrays.asList("B1市_1_2_1", "B1市_1_2_2", "其它"));
                put("B1市_2_1", Arrays.asList("B1市_2_1_1", "B1市_2_1_2", "其它"));
                put("B1市_2_2", Arrays.asList("B1市_2_2_1", "B1市_2_2_2", "其它"));
                put("B2市_1_1", Arrays.asList("B2市_1_1_1", "B2市_1_1_2", "其它"));
            }}
        ));
        final String fileName = "Validation CascadeList Test.xlsx";
        new Workbook()
            .addSheet(new SimpleSheet<>(Collections.singletonList(Arrays.asList("省份", "市区", "市镇", "乡村", "门牌"))).putExtProp(Const.ExtendPropertyKey.DATA_VALIDATION, expectValidations))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            FullSheet sheet = reader.sheet(0).asFullSheet();
            List<Validation> validations = sheet.getValidations();
            assertNotNull(validations);
            assertEquals(validations.size(), 2);
        }
    }

    @Test public void testListValidationOverSize() throws IOException {
        List<String> len255List = Arrays.asList("东山县","南靖县","龙海区","诏安县","龙文区","平和县","漳浦县","芗城区","云霄县","华安县","长泰县","南安市","石狮市","泉港区","金门县","丰泽区","安溪县","永春县","洛江区","德化县","晋江市","鲤城区","惠安县","将乐县","沙县区","永安市","梅列区","尤溪县","建宁县","宁化县","三元区","明溪县","泰宁县","大田县","清流县","古田县","寿宁县","福安市","周宁县","蕉城区","霞浦县","福鼎市","屏南县","柘荣县","同安区","翔安区","思明区","湖里区","集美区","海沧区","仙游县","秀屿区","荔城区","城厢区","涵江区","鼓楼区","平潭县","长乐区","闽清县","福清市","连江县","晋安区","罗源县","闽侯县");
        List<String> over255List = new ArrayList<>(len255List);
        over255List.add("仓山区");
        List<Validation> expectValidations = new ArrayList<>();
        expectValidations.add(new ListValidation<String>().in(len255List).dimension(Dimension.of("A1")).prompt("内联选择框"));
        expectValidations.add(new ListValidation<String>().in(over255List).dimension(Dimension.of("D1")).prompt("引用选择框"));
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
            // 内联选择
            assertEquals(validations.get(0).toString(), expectValidations.get(0).toString());

            // 判断是否为引用
            Validation expectVal2 = new ListValidation<>().dimension(Dimension.of("D1")).prompt("引用选择框").referer(new CrossDimension(sheet2.getName(), new Dimension(1, (short) 1, 1, (short) over255List.size())));
            assertEquals(validations.get(1).toString(), expectVal2.toString());

            // 判断选项内容和顺序
            Iterator<Row> iter = sheet2.iterator();
            assertTrue(iter.hasNext());
            Row row0 = iter.next();
            assertTrue(over255List.size() >= row0.getLastColumnIndex());
            int i = 0;
            for (; i < row0.getLastColumnIndex(); i++) {
                assertEquals(over255List.get(i), row0.getString(i));
            }
            assertEquals(over255List.size(), i);
        }
    }

    @Test public void t() {
        List<String> list = Arrays.asList("A1市_1_1", "A1市_1_2","A1市_2_1", "A1市_2_2","A1市_3_1", "A1市_3_2","A2市_1_1", "A2市_1_2","A2市_2_1", "A2市_2_2","A2市_3_1", "A2市_3_2","A2市_4_1", "A2市_4_2","A2市_5_1", "A2市_5_2","A3区_1_1", "A3区_1_2","A3区_2_1", "A3区_2_2","B1市_1_1", "B1市_1_2","B1市_2_1", "B1市_2_2","B2市_1_1", "B2市_1_2","B2市_2_1", "B2市_2_2","B2市_3_1", "B2市_3_2","B3市_1_1", "B3市_1_2","B3市_2_1", "B3市_2_2","B3市_3_1", "B3市_3_2","B3市_4_1", "B3市_4_2","B4市_1_1", "B4市_1_2","B4市_2_1", "B4市_2_2","D1市_1_1", "D1市_1_2","D1市_2_1", "D1市_2_2","D2市_1_1", "D2市_1_2","D2市_2_1", "D2市_2_2","D2市_3_1", "D2市_3_2","D2市_4_1", "D2市_4_2");
        for (String k : list) {
            System.out.println("put(\""+k+"\", Arrays.asList(\""+k+"_1\", \""+k+"_2\", \"其它\"));");
        }
    }
}

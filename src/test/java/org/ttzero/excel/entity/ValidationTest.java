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
import org.ttzero.excel.entity.e7.XMLWorksheetWriter;
import org.ttzero.excel.reader.Dimension;
import org.ttzero.excel.util.DateUtil;
import org.ttzero.excel.util.StringUtil;

import java.io.IOException;
import java.sql.Time;
import java.text.ParsePosition;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.List;
import java.util.stream.Collectors;

/**
 * @author wangguanquan3 at 2022-08-17 ‏‎20:05:42
 */
public class ValidationTest extends WorkbookTest {
    @Test public void test() throws IOException {
        List<Validation> validations = new ArrayList<>();
        // 下拉框选择“男”，“女”
        validations.add(new ListValidation<>().in("男", "女").dimension(Dimension.of("A1")));
        // B1:E1 单元格只能输入大于1的数
        validations.add(new WholeValidation().greaterThan(1).dimension(Dimension.of("B1:E1")));
        // 限制日期在2022年
        validations.add(new DateValidation().between("2022-01-01", "2022-12-31").dimension(Dimension.of("A2")));
        // 限制时间小于下午6点（因为此时下班...）
        validations.add(new TimeValidation().lessThan(DateUtil.toTimeValue(Time.valueOf("18:00:00"))).dimension(Dimension.of("B2")));
        new Workbook("Validation Test").addSheet(new EmptySheet().putExtProp("dataValidations", validations).setSheetWriter(new ValidationWorksheetWriter())).writeTo(defaultTestPath);
    }

    public static abstract class Validation {
        /**
         * 允许为空
         */
        boolean allowBlank = true;
        /**
         * 显示下拉框
         */
        boolean showInputMessage = true;
        /**
         * 显示提示信息
         */
        boolean showErrorMessage = true;
        /**
         * 作用范围
         */
        Dimension sqref;
        /**
         * 操作符，不指定时默认between
         */
        Operator operator;

        /**
         * 数据校验类型
         */
        public abstract String getType();

        /**
         * 校验内容
         */
        public abstract String validationFormula();

        @Override
        public String toString() {
            return "<dataValidation type=\"" + getType()
                    + (operator == null || operator == Operator.between ? "" : "\" operator=\"" + operator)
                    + "\" allowBlank=\"" + (allowBlank ? 1 : 0)
                    + "\" showInputMessage=\"" + (showInputMessage ? 1 : 0)
                    + "\" showErrorMessage=\"" + (showErrorMessage ? 1 : 0)
                    + "\" sqref=\"" + sqref + "\">"
                    + validationFormula()
                    + "</dataValidation>";
        }

        public enum Operator {
            between, notBetween, equal, notEqual, greaterThan, lessThan, greaterThanOrEqual, lessThanOrEqual
        }

    }

    public static class ListValidation<T> extends Validation {
        List<T> options;
        Dimension referer;

        public ListValidation<T> in(List<T> options) {
            this.options = options;
            return this;
        }

        @SafeVarargs
        public final ListValidation<T> in(T... options) {
            this.options = Arrays.asList(options);
            return this;
        }

        public ListValidation<T> in(Dimension referer) {
            this.referer = referer;
            return this;
        }

        @Override
        public String getType() {
            return "list";
        }

        public ListValidation<T> dimension(Dimension sqref) {
            this.sqref = sqref;
            return this;
        }

        @Override
        public String validationFormula() {
            return "<formula1>\"" + (options != null ? options.stream().map(String::valueOf).collect(Collectors.joining(",")) : referer) + "\"</formula1>";
        }
    }

    public static abstract class Tuple2Validation<V1, V2> extends Validation {
        V1 v1;
        V2 v2;

        public Tuple2Validation<V1, V2> equal(V1 v1) {
            this.operator = Operator.equal;
            this.v1 = v1;
            return this;
        }

        public Tuple2Validation<V1, V2> nowEqual(V1 v1) {
            this.operator = Operator.notEqual;
            this.v1 = v1;
            return this;
        }

        public Tuple2Validation<V1, V2> greaterThan(V1 v1) {
            this.operator = Operator.greaterThan;
            this.v1 = v1;
            return this;
        }

        public Tuple2Validation<V1, V2> greaterThanOrEqual(V1 v1) {
            this.operator = Operator.greaterThanOrEqual;
            this.v1 = v1;
            return this;
        }

        public Tuple2Validation<V1, V2> lessThan(V2 v2) {
            this.operator = Operator.lessThan;
            this.v2 = v2;
            return this;
        }

        public Tuple2Validation<V1, V2> lessThanOrEqual(V2 v2) {
            this.operator = Operator.lessThanOrEqual;
            this.v2 = v2;
            return this;
        }

        public Tuple2Validation<V1, V2> between(V1 v1, V2 v2) {
            this.operator = Operator.between;
            this.v1 = v1;
            this.v2 = v2;
            return this;
        }

        public Tuple2Validation<V1, V2> notBetween(V1 v1, V2 v2) {
            this.operator = Operator.notBetween;
            this.v1 = v1;
            this.v2 = v2;
            return this;
        }

        public Tuple2Validation<V1, V2> dimension(Dimension sqref) {
            this.sqref = sqref;
            return this;
        }

        @Override
        public String validationFormula() {
            String v;
            if (operator == null) operator = Operator.between;
            switch (operator) {
                case equal:
                case notEqual:
                case greaterThan:
                case greaterThanOrEqual:
                    v = "<formula1>" + v1 + "</formula1>";
                    break;
                case lessThan:
                case lessThanOrEqual:
                    v = "<formula1>" + v2 + "</formula1>";
                    break;
                default:
                    v = "<formula1>" + v1 + "</formula1><formula2>" + v2 + "</formula2>";
                    break;
            }
            return v;
        }
    }

    public static class WholeValidation extends Tuple2Validation<Integer, Integer> {
        @Override
        public String getType() {
            return "whole";
        }
    }

    public static class TextLengthValidation extends Tuple2Validation<Integer, Integer> {
        @Override
        public String getType() {
            return "textLength";
        }
    }

    public static class DateValidation extends Tuple2Validation<Integer, Integer> {
        public DateValidation between(Date from, Date to) {
            if (from != null) v1 = DateUtil.toDateValue(from);
            if (to != null) v2 = DateUtil.toDateValue(to);
            return this;
        }

        /**
         * @param from time in format "yyyy-MM-dd"
         * @param to   time in format "yyyy-MM-dd"
         */
        public DateValidation between(String from, String to) {
            if (StringUtil.isNotEmpty(from))
                v1 = DateUtil.toDateValue(DateUtil.dateFormat.get().parse(from.substring(0, Math.min(from.length(), 10)), new ParsePosition(0)));
            if (StringUtil.isNotEmpty(to))
                v2 = DateUtil.toDateValue(DateUtil.dateFormat.get().parse(to.substring(0, Math.min(to.length(), 10)), new ParsePosition(0)));
            return this;
        }

        @Override
        public String getType() {
            return "date";
        }

    }

    public static class TimeValidation extends Tuple2Validation<Double, Double> {
        public TimeValidation between(Date from, Date to) {
            if (from != null) v1 = DateUtil.toTimeValue(from);
            if (to != null) v2 = DateUtil.toTimeValue(to);
            return this;
        }

        public TimeValidation between(Time from, Time to) {
            if (from != null) v1 = DateUtil.toTimeValue(from);
            if (to != null) v2 = DateUtil.toTimeValue(to);
            return this;
        }

        /**
         * @param from time in format "hh:mm:ss"
         * @param to   time in format "hh:mm:ss"
         */
        public TimeValidation between(String from, String to) {
            if (StringUtil.isNotEmpty(from)) v1 = DateUtil.toTimeValue(Time.valueOf(from));
            if (StringUtil.isNotEmpty(to)) v2 = DateUtil.toTimeValue(Time.valueOf(to));
            return this;
        }

        @Override
        public String getType() {
            return "time";
        }
    }


    public static class ValidationWorksheetWriter extends XMLWorksheetWriter {
        @Override
        protected void afterSheetData() throws IOException {
            super.afterSheetData();

            @SuppressWarnings("unchecked")
            List<Validation> validations = (List<Validation>) sheet.getExtPropValue("dataValidations");
            if (validations != null && !validations.isEmpty()) {
                bw.write("<dataValidations count=\"");
                bw.writeInt(validations.size());
                bw.write("\">");
                for (Validation e : validations) {
                    bw.write(e.toString());
                }
                bw.write("</dataValidations>");
            }
        }
    }
}

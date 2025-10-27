/*
 * Copyright (c) 2017-2023, guanquan.wang@hotmail.com All Rights Reserved.
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


package org.ttzero.excel.validation;


import org.dom4j.Element;
import org.ttzero.excel.reader.CrossDimension;
import org.ttzero.excel.reader.Dimension;

/**
 * 范围验证，限定起始和结束值范围
 *
 * @see DateValidation
 * @see TimeValidation
 * @see TextLengthValidation
 * @see WholeValidation
 * @author guanquan.wang at 2022-08-17 20:05:42
 */
public abstract class RangeValidation<T> extends Validation {
    public RangeVal<T> v1, v2;

    public RangeValidation<T> equal(T v1) {
        this.operator = Operator.equal;
        this.v1 = convertToRangeVal(v1);
        return this;
    }

    public RangeValidation<T> nowEqual(T v1) {
        this.operator = Operator.notEqual;
        this.v1 = convertToRangeVal(v1);
        return this;
    }

    public RangeValidation<T> greaterThan(T v1) {
        this.operator = Operator.greaterThan;
        this.v1 = convertToRangeVal(v1);
        return this;
    }

    public RangeValidation<T> greaterThanOrEqual(T v1) {
        this.operator = Operator.greaterThanOrEqual;
        this.v1 = convertToRangeVal(v1);
        return this;
    }

    public RangeValidation<T> lessThan(T v2) {
        this.operator = Operator.lessThan;
        this.v2 = convertToRangeVal(v2);
        return this;
    }

    public RangeValidation<T> lessThanOrEqual(T v2) {
        this.operator = Operator.lessThanOrEqual;
        this.v2 = convertToRangeVal(v2);
        return this;
    }

    public RangeValidation<T> between(T v1, T v2) {
        this.operator = Operator.between;
        this.v1 = convertToRangeVal(v1);
        this.v2 = convertToRangeVal(v2);
        return this;
    }

    public RangeValidation<T> notBetween(T v1, T v2) {
        this.operator = Operator.notBetween;
        this.v1 = convertToRangeVal(v1);
        this.v2 = convertToRangeVal(v2);
        return this;
    }

    static <T> RangeVal<T> convertToRangeVal(T v) {
        RangeVal<T> rv;
        if (v instanceof Dimension) rv = RangeVal.dimensionOf(new CrossDimension((Dimension) v));
        else if (v instanceof CrossDimension) rv = RangeVal.dimensionOf((CrossDimension) v);
        else rv = RangeVal.of(v);
        return rv;
    }

    @Override
    public String validationFormula() {
        boolean ext = isExtension()
            , b1 = v1 != null && v1.referer != null
            , b2 = v2 != null && v2.referer != null;
        if (ext && !(b1 || b2))
            throw new IllegalArgumentException("Extension validation must setting dimension values");
        String v;
        if (operator == null) operator = Operator.between;
        switch (operator) {
            case equal:
            case notEqual:
            case greaterThan:
            case greaterThanOrEqual: v = ext && b1 ? "<x14:formula1><xm:f>" + v1 + " </xm:f></x14:formula1>" : "<formula1>" + v1 + "</formula1>"; break;
            case lessThan:
            case lessThanOrEqual: v = ext && b2 ? "<x14:formula1><xm:f>" + v2 + " </xm:f></x14:formula1>" : "<formula1>" + v2 + "</formula1>"; break;
            default: v = ext ? "<x14:formula1><xm:f>" + v1 + " </xm:f></x14:formula1><x14:formula2><xm:f>" + v2 + " </xm:f></x14:formula2>": "<formula1>" + v1 + "</formula1><formula2>" + v2 + "</formula2>"; break;
        }
        if (ext) v += "<xm:sqref>"+ getSqrefStr() +"</xm:sqref>";
        return v;
    }

    public static class RangeVal<T> {
        /**
         * 引用值
         */
        public CrossDimension referer;
        /**
         * 值
         */
        public T val;

        public static <T> RangeVal<T> dimensionOf(CrossDimension referer) {
            RangeVal<T> rv = new RangeVal<>();
            rv.referer = referer;
            return rv;
        }

        public static <T> RangeVal<T> of(T val) {
            RangeVal<T> rv = new RangeVal<>();
            rv.val = val;
            return rv;
        }

        @Override
        public String toString() {
            return referer != null ? referer.toString() : String.valueOf(val);
        }
    }

    protected abstract T parseTxtValue(String txt);

    @Override
    protected void parseAttribute(Element e, boolean isExt) {
        super.parseAttribute(e, isExt);
        Element formula1 = e.element("formula1"), formula2 = e.element("formula2");
        String txt;
        if (formula1 != null) {
            txt = isExt ? formula1.elementText("f") : formula1.getText();
            this.v1 = parseValue(txt.trim());
        }
        if (formula2 != null) {
            txt = isExt ? formula2.elementText("f") : formula2.getText();
            this.v2 = parseValue(txt.trim());
        }
        // Special case
        if ((operator == Validation.Operator.lessThan || operator == Validation.Operator.lessThanOrEqual) && this.v2 == null) {
            this.v2 = this.v1;
            this.v1 = null;
        }
    }

    protected RangeVal<T> parseValue(String txt) {
        // 0:无 1:字符串 2:坐标 3:跨工作表坐标 4:公式 5:数字
        int type = testValueType(txt);
        if (type == 0) return null;
        RangeVal<T> rv = new RangeVal<>();
        switch (type) {
            case 1:
                rv.val = parseTxtValue(txt.substring(1, txt.length() - 1));
                break;
            case 5:
                rv.val = parseTxtValue(txt);
                break;
            case 2:
            case 3:
                this.referer = CrossDimension.of(txt);
                break;
            default:
                LOGGER.warn("Unsupported formula value[{}]", txt);
        }
        return rv;
    }
}

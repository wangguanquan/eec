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
    public T v1, v2;

    public RangeValidation<T> equal(T v1) {
        this.operator = Operator.equal;
        this.v1 = v1;
        return this;
    }

    public RangeValidation<T> nowEqual(T v1) {
        this.operator = Operator.notEqual;
        this.v1 = v1;
        return this;
    }

    public RangeValidation<T> greaterThan(T v1) {
        this.operator = Operator.greaterThan;
        this.v1 = v1;
        return this;
    }

    public RangeValidation<T> greaterThanOrEqual(T v1) {
        this.operator = Operator.greaterThanOrEqual;
        this.v1 = v1;
        return this;
    }

    public RangeValidation<T> lessThan(T v2) {
        this.operator = Operator.lessThan;
        this.v2 = v2;
        return this;
    }

    public RangeValidation<T> lessThanOrEqual(T v2) {
        this.operator = Operator.lessThanOrEqual;
        this.v2 = v2;
        return this;
    }

    public RangeValidation<T> between(T v1, T v2) {
        this.operator = Operator.between;
        this.v1 = v1;
        this.v2 = v2;
        return this;
    }

    public RangeValidation<T> notBetween(T v1, T v2) {
        this.operator = Operator.notBetween;
        this.v1 = v1;
        this.v2 = v2;
        return this;
    }

    @Override
    public String validationFormula() {
        boolean ext = isExtension()
            , b1 = v1 != null && v1 instanceof Dimension
            , b2 = v2 != null && v2 instanceof  Dimension;
        if (ext && !(b1 || b2))
            throw new IllegalArgumentException("Extension validation must setting dimension values");
        String v;
        if (operator == null) operator = Operator.between;
        switch (operator) {
            case equal:
            case notEqual:
            case greaterThan:
            case greaterThanOrEqual: v = ext && b1 ? "<x14:formula1><xm:f>" + otherSheetName + ":" + ((Dimension) v1).toReferer() + " </xm:f></x14:formula1>" : "<formula1>" + v1 + "</formula1>"; break;
            case lessThan:
            case lessThanOrEqual: v = ext && b2 ? "<x14:formula1><xm:f>" + otherSheetName + ":" + ((Dimension) v2).toReferer() + " </xm:f></x14:formula1>" : "<formula1>" + v2 + "</formula1>"; break;
            default: v = ext ? "<x14:formula1><xm:f>" + otherSheetName + ":" + ((Dimension) v1).toReferer() + " </xm:f></x14:formula1><x14:formula2><xm:f>" + otherSheetName + ":" + ((Dimension) v2).toReferer() + " </xm:f></x14:formula2>": "<formula1>" + v1 + "</formula1><formula2>" + v2 + "</formula2>"; break;
        }
        if (ext) v += "<xm:sqref>"+ getSqrefStr() +"</xm:sqref>";
        return v;
    }
}

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


package org.ttzero.excel.validation;


/**
 * 范围验证，限定起始和结束值范围
 *
 * @see DateValidation
 * @see TimeValidation
 * @see TextLengthValidation
 * @see WholeValidation
 * @author guanquan.wang at 2022-08-17 20:05:42
 */
public abstract class Tuple2Validation<V1, V2> extends Validation {
    public V1 v1;
    public V2 v2;

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

    @Override
    public String validationFormula() {
        String v;
        if (operator == null) operator = Operator.between;
        switch (operator) {
            case equal:
            case notEqual:
            case greaterThan:
            case greaterThanOrEqual: v = "<formula1>" + v1 + "</formula1>"; break;
            case lessThan:
            case lessThanOrEqual: v = "<formula1>" + v2 + "</formula1>"; break;
            default: v = "<formula1>" + v1 + "</formula1><formula2>" + v2 + "</formula2>"; break;
        }
        return v;
    }
}

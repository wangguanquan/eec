/*
 * Copyright (c) 2017-2022, guanquan.wang@yandex.com All Rights Reserved.
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


package org.ttzero.excel.annotation;

import java.math.BigDecimal;

/**
 * @author guanquan.wang at 2022-05-03 18:38
 */
public class CellData<T> {

    private BigDecimal numberValue;

    private String stringValue;

    private Boolean booleanValue;

    /**
     * The resulting converted data.
     */
    private T data;

    public CellData(CellData<T> other) {
        this.numberValue = other.numberValue;
        this.stringValue = other.stringValue;
        this.booleanValue = other.booleanValue;
        this.data = other.data;
    }

    public CellData() {}

    public CellData(T data) {
        this.data = data;
    }

    public BigDecimal getNumberValue() {
        return numberValue;
    }

    public void setNumberValue(BigDecimal numberValue) {
        this.numberValue = numberValue;
    }

    public String getStringValue() {
        return stringValue;
    }

    public void setStringValue(String stringValue) {
        this.stringValue = stringValue;
    }

    public Boolean getBooleanValue() {
        return booleanValue;
    }

    public void setBooleanValue(Boolean booleanValue) {
        this.booleanValue = booleanValue;
    }

    public T getData() {
        return data;
    }

    public void setData(T data) {
        this.data = data;
    }


    @Override
    public String toString() {
       if (stringValue != null) {
           return stringValue;
       }
       if (numberValue != null) {
           return numberValue.toString();
       }
       if (booleanValue != null) {
           return booleanValue.toString();
       }
       return data.toString();
    }

}

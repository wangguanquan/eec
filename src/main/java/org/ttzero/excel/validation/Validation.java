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
 * 数据验证
 *
 * @author guanquan.wang at 2022-08-17 20:05:42
 */
public abstract class Validation {
    /**
     * 允许为空
     */
    public boolean allowBlank = true;
    /**
     * 显示下拉框
     */
    public boolean showInputMessage = true;
    /**
     * 显示提示信息
     */
    public boolean showErrorMessage = true;
    /**
     * 作用范围
     */
    public Dimension sqref;
    /**
     * 操作符，不指定时默认between
     */
    public Operator operator;

    /**
     * 数据校验类型
     *
     * @return 数据验类型，包含 {@code list}序列, {@code whole}整数, {@code date}日期, {@code time}时间, {@code textLength}文本长度
     */
    public abstract String getType();

    /**
     * 校验内容
     *
     * @return 验证对象转xml文本
     */
    public abstract String validationFormula();

    /**
     * 是否为扩展节点
     *
     * @return {@code true} 扩展节点
     */
    public boolean isExtension() {
        return false;
    }

    public Validation dimension(Dimension sqref) {
        this.sqref = sqref;
        return this;
    }

    @Override
    public String toString() {
        return "<" + (isExtension() ? "x14:" : "" ) + "dataValidation type=\"" + getType()
            + (operator == null || operator == Operator.between ? "" : "\" operator=\"" + operator)
            + "\" allowBlank=\"" + (allowBlank ? 1 : 0)
            + "\" showInputMessage=\"" + (showInputMessage ? 1 : 0)
            + "\" showErrorMessage=\"" + (showErrorMessage ? 1 : 0)
            + (isExtension() ? "\" xr:uid=\"{E9742D38-9313-3C47-9945-211275B11887}\">" : "\" sqref=\"" + sqref + "\">")
            + validationFormula()
            + "</" + (isExtension() ? "x14:" : "" ) + "dataValidation>";
    }

    public enum Operator {
        between, notBetween, equal, notEqual, greaterThan, lessThan, greaterThanOrEqual, lessThanOrEqual
    }
}

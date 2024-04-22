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

import java.util.Arrays;
import java.util.List;
import java.util.stream.Collectors;

/**
 * 序列验证，限定单元格的值只能在序列中选择
 *
 * @author guanquan.wang at 2022-08-17 20:05:42
 */
public class ListValidation<T> extends Validation {
    /**
     * 可选的序列值
     */
    public List<T> options;
    /**
     * 引用其它可选序列的坐标
     */
    public Dimension referer;

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

    @Override
    public String validationFormula() {
        return "<formula1>\"" + (options != null ? options.stream().map(String::valueOf).collect(Collectors.joining(",")) : referer) + "\"</formula1>";
    }
}

/*
 * Copyright (c) 2017-2022, guanquan.wang@yandex.com All Rights Reserved.
 * and/or licensed to one or more contributor license agreements.
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

import org.ttzero.excel.processor.StyleProcessor;

import java.lang.annotation.ElementType;
import java.lang.annotation.Inherited;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 动态样式处理器
 *
 * <p>动态样式是指根据行数据为每个单元格或整行设置不同样式，这个功能可以极大丰富文件的可读性和多样性.
 * StyleDesign注解作用于{@code type}类时修改整行样式，作用于{@code field}和{@code method}
 *  时修改单个单元格样式。StyleDesign指定的样式处理器必须实现{@link StyleProcessor}接口，</p>
 *
 * <p>参考文档:</p>
 * <p><a href="https://github.com/wangguanquan/eec/wiki/4-%E5%8A%A8%E6%80%81%E8%AE%BE%E7%BD%AE%E6%A0%B7%E5%BC%8F">动态设置样式</a></p>
 *
 * @see StyleProcessor
 * @author suyl at 2022-03-23 17:38
 *
 */
@Target({ElementType.TYPE, ElementType.FIELD, ElementType.METHOD})
@Retention(RetentionPolicy.RUNTIME)
@Inherited
public @interface StyleDesign {
    /**
     * 指定动态样式处理器
     *
     * @return 样式处理器
     */
    Class<? extends StyleProcessor> using() default StyleProcessor.None.class;
}


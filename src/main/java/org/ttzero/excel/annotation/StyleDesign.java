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
 * Customize Style design
 *
 * @author suyl at 2022-03-23 17:38
 *
 */
@Target(ElementType.TYPE)
@Retention(RetentionPolicy.RUNTIME)
@Inherited
public @interface StyleDesign {
    /**
     * Specify a {@link StyleProcessor} to setting the cell style
     *
     * @return a {@link StyleProcessor} class
     */
    Class<? extends StyleProcessor> using() default StyleProcessor.None.class;
}


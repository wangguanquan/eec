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

import java.lang.annotation.Documented;
import java.lang.annotation.ElementType;
import java.lang.annotation.Inherited;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * If you have a large table of data in Excel, it can be useful to
 * freeze rows or columns. This way you can keep rows or columns
 * visible while scrolling through the rest of the worksheet.
 * <p>
 *
 *
 * @author guanquan.wang at 2022-04-17 11:35
 */
@Target(ElementType.TYPE)
@Retention(RetentionPolicy.RUNTIME)
@Inherited
@Documented
public @interface FreezePanes {
    /**
     * Specify the top row to freeze(one-base). Negative numbers are not allowed
     *
     * @return the top row number to freeze, 0 means unfreeze
     */
    int topRow() default 0;

    /**
     * Specify the first column to freeze(one-base), Negative numbers are not allowed
     *
     * @return the first column number to freeze, 0 means unfreeze
     */
    int firstColumn() default 0;
}

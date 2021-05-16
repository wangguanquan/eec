/*
 * Copyright (c) 2017-2020, guanquan.wang@yandex.com All Rights Reserved.
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

import org.ttzero.excel.entity.style.Fill;
import org.ttzero.excel.entity.style.Font;

import java.lang.annotation.Documented;
import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Custom header styles
 *
 * @author jialei2 at 021-05-10 17:38
 */
@Target({ElementType.FIELD, ElementType.METHOD, ElementType.TYPE})
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface HeaderStyle {

    /**
     * The {@link Font} property is a shorthand property for:
     * <ul>
     * <li>font-style: Specifies the font {@link Font.Style}. Default value is "normal"</li>
     * <li>font-size: Specifies the font size, Default value is 12</li>
     * <li>font-family: Specifies the font family. Default value depends on the property {@code local-font-family}</li>
     * <li>color: Specifies the font {@link java.awt.Color}. Default value is {@link java.awt.Color#BLACK}</li>
     * </ul>
     *
     * @see Font#parse(String)
     * @return all properties join with {@code ' '} or {@code '_'}
     */
    String font() default "bold 12 black";

    /**
     * The {@link Fill} property is a shorthand property for:
     * <ul>
     * <li>fg-color: Specifies the foreground color. Default value is "#666699"</li>
     * <li>bg-color: Specifies the background color. Default value is "#666699"</li>
     * <li>pattern-type</li>
     * </ul>
     *
     * @see Fill#parse(String)
     * @return all properties join with {@code ' '}
     */
    String fill() default "#666699 solid";

    /**
     * The {@link Fill} property is a shorthand property for:
     * <ul>
     * <li>fgColor</li>
     * <li>bgColor</li>
     * <li>patternType</li>
     * </ul>
     *
     * @see Fill#parse(String)
     * @return all properties join with {@code ' '}
     */
    String border() default "thin black";

}

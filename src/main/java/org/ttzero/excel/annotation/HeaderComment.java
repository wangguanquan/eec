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

import java.lang.annotation.Documented;
import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 表头批注，{@code title}将加粗显示但不是必须的，大多数情况下将批注文本放到{@code value}下即可，
 * 描述文本较多时需要使用{@code width}和{@code height}两个属性来调整弹出框的大小以将内容显示完全
 *
 * @author guanquan.wang at 2020-05-21 16:43
 */
@Target({ElementType.FIELD, ElementType.METHOD})
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface HeaderComment {
    /**
     * 批注正文
     *
     * @return 正文文本，为空时不显示批注
     */
    String value() default "";

    /**
     * 批注标题，加粗显示
     *
     * @return 标题，可为空
     */
    String title() default "";

    /**
     * 指定批注弹出框宽度
     *
     * @return 批注宽度
     */
    double width() default 100.8D;

    /**
     * 指定批注弹出框高度
     *
     * @return 批注高度
     */
    double height() default 60.6D;
}

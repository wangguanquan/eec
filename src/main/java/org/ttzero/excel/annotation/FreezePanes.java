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
 * 设置行列“冻结”，滚动工作表时保持首行或首列总是可见，此功能为增强型功能。
 *
 * <p>注意：不能指定一个范围值，冻结必须从首行和首列开始，只能指定结尾行/列号，如{@code FreezePanes(topRow = 3)}
 * 这段代码将冻结前3行即1,2,3这三行在滚动工作表时总是可见且总是在顶部，{@code firstColumn}也是同样的效果</p>
 *
 * @author guanquan.wang at 2022-04-17 11:35
 */
@Target(ElementType.TYPE)
@Retention(RetentionPolicy.RUNTIME)
@Inherited
@Documented
public @interface FreezePanes {
    /**
     * 指定冻结的结尾行号（从1开始），0和负数表示不冻结
     *
     * @return 从1开始的行号，0和负数表示不冻结
     */
    int topRow() default 0;

    /**
     * 指定冻结的结尾列号（从1开始），0和负数表示不冻结
     *
     * @return 从1开始的列号，0和负数表示不冻结
     */
    int firstColumn() default 0;
}

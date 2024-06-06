/*
 * Copyright (c) 2017-2021, guanquan.wang@yandex.com All Rights Reserved.
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

package org.ttzero.excel.processor;

/**
 * 输出转换器，将Java数据类型转为Excel输出类型，通常用于将状态值、枚举值转换为文本
 *
 * <pre>
 * new Workbook()
 *   .addSheet(new ListSheet&lt;&gt;(
 *      new Column("审核单号", "billNo"),
 *      new Column("审核状态", "status", n -&gt; AuditStatus.byVal((int)n).desc())
 *   ))
 *   .writeTo(Paths.get("/tmp/"));</pre>
 *
 * @author guanquan.wang on 2021-11-30 19:10
 */
@FunctionalInterface
public interface ConversionProcessor {
    /**
     * 输出转换器，导出Excel时将数据转换后输出
     *
     * @param v 原始值
     * @return 转换后的值
     */
    Object conversion(Object v);
}

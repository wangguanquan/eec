/*
 * Copyright (c) 2009, guanquan.wang@yandex.com All Rights Reserved.
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

package cn.ttzero.excel.processor;

import cn.ttzero.excel.entity.style.Styles;

/**
 * 样式转换器
 * Created by guanquan.wang on 2017/10/13.
 */
@FunctionalInterface
public interface StyleProcessor {
    /**
     * 样式转换器
     * 添加样式时必须使用sst.add方法添加，然后将返回的int值做为转换器的返回值
     * eg:
     * <pre><code lang='java'>
     *    StyleProcessor sp = (o, style, sst) // 将背景改为黄色
     *      -> style |= Styles.clearFill(style) | sst.addFill(new Fill(Color.yellow));
     * </code></pre>
     * @param o 当前单元格值
     * @param style 当前单元格样式
     * @param sst 样式类，整个Workbook共享样式
     * @return 新样式
     */
    int build(Object o, int style, Styles sst);
}

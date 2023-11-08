/*
 * Copyright (c) 2017, guanquan.wang@yandex.com All Rights Reserved.
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

import org.ttzero.excel.entity.style.Styles;


/**
 * 动态样式处理器，根据行数据动态修改样式，可以非常简单的实现高亮效果，
 * 对于预警类型的导出尤为实用。
 *
 * <p>处理器有3个重要的参数：</p>
 *  <ol>
 *    <li>o: 单元格的值，作用于行级时o为一个Bean实体</li>
 *    <li>style: 当前单元格样式值</li>
 *    <li>sst: 全局的样式对象{@link Styles}</li>
 *  </ol>
 *
 * <p>{@code StyleProcessor}可作用于行级或者单元格，放在工作表上可修改整行样式，
 * 放在单个{@code Column}上作用于单个单元格，</p>
 *
 * @author guanquan.wang on 2017/10/13.
 */
@FunctionalInterface
public interface StyleProcessor<T> {
    /**
     * 动态样式处理，修改样式请使用{@code Styles.modifyXX}方法
     *
     * <blockquote><pre>
     * StyleProcessor sp = (o, style, sst)
     *     // 库存小于100时高亮显示 - 填充黄色
     *     -&gt; o &lt; 100 ? style |= sst.modifyFill(new Fill(Color.yellow)) : style;
     * </pre></blockquote>
     *
     * @param o     单元格的值
     * @param style 当前单元格样式值
     * @param sst   全局的样式对象{@link Styles}
     * @return 新的样式值
     */
    int build(T o, int style, Styles sst);

    /**
     * 无动态样式，默认
     */
    final class None implements StyleProcessor<Object> {

        @Override
        public final int build(Object o, int style, Styles sst) {
            return style;
        }
    }
}

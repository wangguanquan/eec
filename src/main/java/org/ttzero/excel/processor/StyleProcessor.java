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

import org.ttzero.excel.entity.Axis;
import org.ttzero.excel.entity.style.Styles;


/**
 * The style conversion
 *
 * @author guanquan.wang on 2017/10/13.
 */
@FunctionalInterface
public interface StyleProcessor<T> {
    /**
     * The style conversion
     * You must add it using {@code Styles#addXXX} method before adding a style,
     * and then use the returned int value as the return value of the converter.
     * <blockquote><pre>
     * StyleProcessor sp = (o, style, sst) // Fill of 'yellow' color
     *     -&gt; style |= Styles.clearFill(style) | sst.addFill(new Fill(Color.yellow));
     * </pre></blockquote>
     *
     * @param o     the value of cell
     * @param style the current style of cell
     * @param sst   the {@link Styles} entry
     * @param axis  the axis of cell
     * @return new style of cell
     */
    int build(T o, int style, Styles sst, Axis axis);

    /**
     * None processor
     */
    final class None implements StyleProcessor<Object> {

        @Override
        public int build(Object o, int style, Styles sst, Axis axis) {
            return style;
        }
    }
}

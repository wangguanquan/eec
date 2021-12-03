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
 * A value Conversion
 * Typically used to convert state values or enumerated values
 * to meaningful values
 *
 * @author guanquan.wang on 2021-11-30 19:10
 */
@FunctionalInterface
public interface ConversionProcessor {
    /**
     * A value Converter, the converted value is used as the export value
     * and the style is also changed accordingly
     *
     * @param v the original value
     * @return the converted value
     */
    Object conversion(Object v);
}

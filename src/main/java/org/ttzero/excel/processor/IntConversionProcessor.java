/*
 * Copyright (c) 2019, guanquan.wang@yandex.com All Rights Reserved.
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
 * An Integer Conversion
 * Typically used to convert state values or enumerated values
 * to meaningful values
 * <p>
 * Created by guanquan.wang on 2017/10/13.
 */
@FunctionalInterface
public interface IntConversionProcessor {
    /**
     * The integer value include byte, char, short, int
     *
     * @param n the integer value
     * @return the converted value
     */
    Object conversion(int n);
}

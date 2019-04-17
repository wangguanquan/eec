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

package cn.ttzero.excel.processor;

/**
 * Int值转其它有意义值
 * 一般用于将状态值，或者枚举值转换为用户感知的具有实际意义的值
 * Created by guanquan.wang on 2017/10/13.
 */
@FunctionalInterface
public interface IntConversionProcessor {
    /**
     * Int值包括byte, char, short, int
     * @param n 数据库值或原对象值
     * @return 转换后的值
     */
    Object conversion(int n);
}

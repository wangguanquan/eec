/*
 * Copyright (c) 2017-2023, guanquan.wang@yandex.com All Rights Reserved.
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


package org.ttzero.excel.reader;


/**
 * 列属性
 *
 * @author guanquan.wang at 2023-12-07 14:08
 */
public class Col {
    /**
     * 列索引范围(one base)
     */
    public int min, max;
    /**
     * 列宽
     */
    public double width;
    /**
     * 是否隐藏列
     */
    public boolean hidden;

    public Col() { }
    public Col(int min, int max, double width) {
        this.min = min;
        this.max = max;
        this.width = width;
    }

    public Col(int min, int max, double width, boolean hidden) {
        this(min, max, width);
        this.hidden = hidden;
    }

    @Override
    public String toString() {
        return "min: " + min + ", max:" + max + ", width:" + width + ", hidden:" + hidden;
    }
}

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

package org.ttzero.excel.entity.style;

/**
 * 定义边框样式
 * <p>
 * Created by guanquan.wang on 2018-02-06 08:55
 */
public enum BorderStyle {
    NONE("none"),
    THIN("thin"),    // 细直线
    MEDIUM("medium"),
    DASHED("dashed"), // 虚线
    DOTTED("dotted"), // 点线
    THICK("thick"), // 粗直线
    DOUBLE("double"), // 双线
    HAIR("hair"),
    MEDIUM_DASHED("mediumDashed"),
    DASH_DOT("dashDot"),
    MEDIUM_DASH_DOT("mediumDashDot"),
    DASH_DOT_DOT("dashDotDot"),
    MEDIUM_DASH_DOT_DOT("mediumDashDotDot"),
    SLANTED_DASH_DOT("slantDashDot");

    private String name;

    BorderStyle(String name) {
        this.name = name;
    }

    public String getName() {
        return name;
    }

    public static BorderStyle getByName(String name) {
        BorderStyle[] borderStyles = values();
        for (BorderStyle borderStyle : borderStyles) {
            if (borderStyle.name.equals(name)) {
                return borderStyle;
            }
        }
        return null;
    }
}

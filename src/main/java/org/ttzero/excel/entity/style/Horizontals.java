/*
 * Copyright (c) 2017-2018, guanquan.wang@yandex.com All Rights Reserved.
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
 * @author guanquan.wang at 2018-02-11 15:02
 */
public class Horizontals {
    // General Horizontal Alignment( Text data is left-aligned. Numbers
    // , dates, and times are right-aligned.Boolean types are centered)
    public static final int GENERAL = 0
        , LEFT = 1 << Styles.INDEX_HORIZONTAL // Left Horizontal Alignment
        , RIGHT = 2 << Styles.INDEX_HORIZONTAL // Right Horizontal Alignment
        , CENTER = 3 << Styles.INDEX_HORIZONTAL // Centered Horizontal Alignment
        , CENTER_CONTINUOUS = 4 << Styles.INDEX_HORIZONTAL // (Center Continuous Horizontal Alignment
        , FILL = 5 << Styles.INDEX_HORIZONTAL // Fill
        , JUSTIFY = 6 << Styles.INDEX_HORIZONTAL // Justify
        , DISTRIBUTED = 7 << Styles.INDEX_HORIZONTAL // Distributed Horizontal Alignment
        ;

    private static final String[] _names = {
            "general"
            ,"left"
            ,"right"
            ,"center"
            ,"centerContinuous"
            ,"fill"
            ,"justify"
            ,"distributed"
    };

    public static String of(int n) {
        return _names[n];
    }
}

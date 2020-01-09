/*
 * Copyright (c) 2017-2019, guanquan.wang@yandex.com All Rights Reserved.
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

package org.ttzero.excel.manager;

/**
 * Type of excel. Biff8 or XLSX(Office open xml)
 *
 * @author guanquan.wang at 2019-01-24 10:12
 */
public enum ExcelType {
    /**
     * BIFF8 only
     * Excel 8.0 Excel 97
     * Excel 9.0 Excel 2000
     * Excel 10.0 Excel XP
     * Excel 11.0 Excel 2003
     */
    XLS,

    /**
     * Excel 12.0(2007) or later
     */
    XLSX,

    /**
     * Others
     */
    UNKNOWN
}

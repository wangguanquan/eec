/*
 * Copyright (c) 2019-2020, guanquan.wang@yandex.com All Rights Reserved.
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
 * Create by guanquan.wang at 2019-12-20 10:07
 */
public class Dimension {
    public final int firstRow;
    public final int lastRow;
    public final short firstColumn;
    public final short lastColumn;

    Dimension(int firstRow, int lastRow, short firstColumn, short lastColumn) {
        this.firstRow = firstRow;
        this.lastRow = lastRow;
        this.firstColumn = firstColumn;
        this.lastColumn = lastColumn;
    }

    @Override
    public String toString() {
        return "{\"first-row\": " + firstRow + ", \"last-row\": " + lastRow
            + ", \"first-column\": " + firstColumn + ", \"last-column\": " + lastColumn + "}";
    }
}

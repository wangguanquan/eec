/*
 * Copyright (c) 2022, guanquan.wang@yandex.com All Rights Reserved.
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


package org.ttzero.excel.entity;

import static org.ttzero.excel.entity.Sheet.int2Col;

/**
 * Axis of the cell(Zero-base)
 *
 * @author guanquan.wang at 2022-09-15 19:22
 */
public class Axis {
    public int row, col;

    public Axis() { }

    public Axis(int row, int col) {
        this.row = row;
        this.col = col;
    }

    public Axis reset(int row, int col) {
        this.row = row;
        this.col = col;
        return this;
    }

    public int getRow() {
        return row;
    }

    public int getCol() {
        return col;
    }

    @Override
    public String toString() {
        return new String(int2Col(col + 1)) + (row + 1);
    }
}

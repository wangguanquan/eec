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

package cn.ttzero.excel.entity;

import cn.ttzero.excel.reader.Cell;

/**
 * Create by guanquan.wang at 2019-04-23 09:57
 */
public class Row {
    // Index to row
    int index = -1;
    // Index to first column (zero base)
    int fc = 0;
    // Index to last column (zero base)
    int lc = -1;
    // Index to XF record
    int xf;
    // Share cell
    Cell[] cells;

    public int getIndex() {
        return index;
    }

    public int getFc() {
        return fc;
    }

    public int getLc() {
        return lc;
    }

    public Cell[] getCells() {
        return cells;
    }

    /**
     * Malloc
     * @param n size_t
     */
    public Cell[] malloc(int n) {
        return cells = new Cell[n];
    }

    /**
     * Malloc and clear
     * @param n size_t
     */
    public Cell[] calloc(int n) {
        malloc(n);
        for (int i = 0; i < n; i++) {
            cells[i] = new Cell();
        }
        return cells;
    }

    /**
     * Resize and clear
     * @param n size_t
     */
    public Cell[] realloc(int n) {
        if (cells == null || cells.length < n) {
            calloc(n);
        }
        return cells;
    }
}

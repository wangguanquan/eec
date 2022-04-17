/*
 * Copyright (c) 2017-2022, guanquan.wang@yandex.com All Rights Reserved.
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

/**
 * Panes
 *
 * @author guanquan.wang at 2022-04-17 10:38
 */
public class Panes {
    /**
     * Panes row
     */
    public int row;
    /**
     * Panes col
     */
    public int col;

    public Panes() { }

    public Panes(int row, int col) {
        this.row = row;
        this.col = col;
    }

    public static Panes of(final int row, final int col) {
        return new Panes(row, col);
    }

    public static Panes row(final int row) {
        return new Panes(row, 0);
    }

    public static Panes col(final int col) {
        return new Panes(0, 2);
    }

    public int getRow() {
        return row;
    }

    public void setRow(int row) {
        this.row = row;
    }

    public int getCol() {
        return col;
    }

    public void setCol(int col) {
        this.col = col;
    }
}

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

import java.util.Iterator;

import static cn.ttzero.excel.manager.Const.ROW_BLOCK_SIZE;

/**
 * All cells in the Cell Table are divided into blocks of 32 consecutive rows, called Row Blocks.
 * The first Row Block starts with the first used row in that sheet.
 * Inside each Row Block there will occur ROW records describing
 * the properties of the rows, and cell records with all the cell contents in this Row Block
 * <p>
 * Create by guanquan.wang at 2019-04-23 08:50
 */
public class RowBlock implements Iterator<Row> {
    private Row[] rows;
    private int i, n, total = 0;
    private boolean eof;
    private int limit;

    public RowBlock() {
        this(ROW_BLOCK_SIZE);
    }

    public RowBlock(int limit) {
        this.limit = limit;
        init();
    }

    private void init() {
        rows = new Row[limit];
        for (int i = 0; i < limit; i++) {
            rows[i] = new Row();
        }
    }

    /**
     * re-open the row-block
     *
     * @return self
     */
    public final RowBlock reopen() {
        eof = false;
//        total = 0;
        return this;
    }

    /**
     * Clear index mark
     */
    public final RowBlock clear() {
        i = n = 0;
        return this;
    }

    /**
     * Total rows of a worksheet
     *
     * @return the total rows
     */
    public int getTotal() {
        return total;
    }

    /**
     * End of file mark
     */
    private void markEnd() {
        eof = true;
    }

    /**
     * End of file mark
     *
     * @return true if end of file
     */
    public boolean isEof() {
        return eof;
    }

    final RowBlock flip() {
        if (i < limit) {
            markEnd();
        }
        n = i;
        total += i;
        i = 0;
        return this;
    }

    public boolean hasNext() {
        return i < n;
    }

    public Row next() {
        return rows[i++];
    }

    public Row firstRow() {
        return rows[0];
    }

    public Row lastRow() {
        return rows[n - 1];
    }

    public int size() {
        return n;
    }

}

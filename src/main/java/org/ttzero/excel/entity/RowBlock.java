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

package org.ttzero.excel.entity;

import java.util.Iterator;

import static org.ttzero.excel.manager.Const.ROW_BLOCK_SIZE;

/**
 * {@code RowBlock}行块由多个连续的{@link Row}行组成，默认包含连续的{@code 32}个{@link Row}行。
 *
 * <p>行块被设计为一个滑行窗口，它类似于{@link java.nio.Buffer}和{@link java.util.Iterator}的合体，
 * 写数据之前调用{@link #clear}方法使游标复原，写满块数据后调用{@link #flip}方法切换为读模式，
 * 这点与{@code Buffer}有着相似特性，读取时使用迭代模式。</p>
 *
 * <p>注意：本类的所有取数方法均限制在当前批次下，不能跨批次获取数据</p>
 *
 * @author guanquan.wang at 2019-04-23 08:50
 */
public class RowBlock implements Iterator<Row> {
    /**
     * 内存连续的一组行数据
     */
    private Row[] rows;
    /**
     * i与Buffer中position等价
     * n与Buffer中limit等价
     * total记录行填共装填了多少行数据
     */
    private int i, n, total = 0;
    /**
     * 当本次装填数据少于块容量时被标记为{@code EOF}
     */
    private boolean eof;
    /**
     * 与Buffer中capacity等价
     */
    private final int limit;

    /**
     * 以默认大小实例化行块，默认{@code 32}
     */
    public RowBlock() {
        this(ROW_BLOCK_SIZE);
    }

    /**
     * 实例化行块并指定容量
     *
     * @param limit 容量
     */
    public RowBlock(int limit) {
        this.limit = limit;
        init();
    }

    /**
     * 创建连续行共享区并实例化行对象
     */
    private void init() {
        rows = new Row[limit];
        for (int i = 0; i < limit; i++) {
            rows[i] = new Row();
        }
    }

    /**
     * 重新打开行块，{@code reopen}将标记清除以达到重用的目的
     *
     * @return 当前行块
     */
    public final RowBlock reopen() {
        eof = false;
        total = 0;
        return this;
    }

    /**
     * 游标复原
     *
     * @return 当前行块
     */
    public final RowBlock clear() {
        i = n = 0;
        // 清除行属性
        for (Row row : rows) {
            row.height = null;
            row.hidden = false;
            row.outlineLevel = null;
        }
        return this;
    }

    /**
     * 获取行块共装填了多少数据，{@link #reopen}方法可清除此记录
     *
     * @return 行块共装填的数据个数
     */
    public int getTotal() {
        return total;
    }

    /**
     * 标记行块已结束，后续将不再装填数据
     */
    private void markEnd() {
        eof = true;
    }

    /**
     * 是否已结束
     *
     * @return true 已结束
     */
    public boolean isEOF() {
        return eof;
    }

    /**
     * 切换为读模式
     *
     * @return 当前行块
     */
    public final RowBlock flip() {
        if (i < limit) {
            markEnd();
        }
        n = i;
        total += i;
        i = 0;
        return this;
    }

    /**
     * 获取容器大小
     *
     * @return 行块容器大小
     */
    public final int capacity() {
        return limit;
    }

    /**
     * 判断迭代器是否更多数据
     *
     * @return Row
     */
    @Override
    public boolean hasNext() {
        return i < n;
    }

    /**
     * 迭代取数
     *
     * @return Row
     */
    @Override
    public Row next() {
        return rows[i++];
    }

    /**
     * 获取本批次行块中第一个数据
     *
     * @return Row
     */
    public Row firstRow() {
        return rows[0];
    }

    /**
     * 获取本批次行块中最后一个数据
     *
     * @return Row
     */
    public Row lastRow() {
        Row row;
        if (n >= 1) row = rows[n - 1];
        else {
            int i = 0;
            for (int len = rows.length - 1; i < len; i++) {
                if (rows[i] == null || rows[i].index >= rows[i + 1].index) {
                    break;
                }
            }
            row = rows[i];
        }
        return row;
    }

    /**
     * 获取本批次指定游标的Row，此方法不会修改游标位置
     *
     * @param position 游标
     * @return Row
     */
    public Row get(int position) {
        return rows[position];
    }

    /**
     * 本批次共装填了多少数据
     *
     * @return 本批次装填个数
     */
    public int size() {
        return n;
    }

    /**
     * 设置游标到指定位置
     *
     * @param position 指定下标
     */
    public void position(int position) {
        if (position < 0 || position >= n)
            throw new ArrayIndexOutOfBoundsException("Index: " + position + ", Size: " + n);
        i = position;
    }

    /**
     * 获取当前游标
     *
     * @return 当前洲标位置
     */
    public int position() {
        return i;
    }
}

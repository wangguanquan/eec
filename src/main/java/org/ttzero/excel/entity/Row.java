/*
 * Copyright (c) 2017-2019, guanquan.wang@hotmail.com All Rights Reserved.
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

import org.ttzero.excel.reader.Cell;

/**
 * 行数据，每个{@code Row}对象包含一组连续的{@link Cell}单元格，
 * 它的设计与在Office中看到的结构完全一样
 *
 * @author guanquan.wang at 2019-04-23 09:57
 */
public class Row {
    // Index to row
    public int index = -1;
    // Index to first column (zero base)
    public int fc = 0;
    // Index to last column (zero base)
    public int lc = -1;
    // Share cell
    public Cell[] cells;
    // height
    public Double height;
    // Is row hidden
    public boolean hidden;
    // Outline level
    public Integer outlineLevel;

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
     * 分配指定大小的连续单元格
     *
     * @param n 单元格数量
     * @return 单元格数组
     */
    public Cell[] malloc(int n) {
        return cells = new Cell[lc = n];
    }

    /**
     * 分配指定大小的连续单元格并初始化
     *
     * @param n 单元格数量
     * @return 单元格数组
     */
    public Cell[] calloc(int n) {
        malloc(n);
        for (int i = 0; i < n; i++) {
            cells[i] = new Cell(i);
        }
        return cells;
    }

    /**
     * 比较并重分配连续{@code n}个单元格，此方法会比较传入的参数{@code n}与当前单元格数量比较，
     * 当{@code n}大于当前数量时才进行重分配
     *
     * @param n 单元格数量
     * @return 单元格数组
     */
    public Cell[] realloc(int n) {
        if (cells == null || cells.length < n) calloc(n);
        lc = n;
        return cells;
    }

    /**
     * 获取行高
     *
     * @return 行高
     */
    public Double getHeight() {
        return height;
    }

    /**
     * 设置行高
     *
     * @param height 行高
     * @return 当前行
     */
    public Row setHeight(Double height) {
        this.height = height;
        return this;
    }

    /**
     * 判断当前行是否隐藏
     *
     * @return true: 隐藏
     */
    public boolean isHidden() {
        return hidden;
    }

    /**
     * 设置当前行显示或隐藏
     *
     * @param hidden true：隐藏当前行 false: 显示
     */
    public void setHidden(boolean hidden) {
        this.hidden = hidden;
    }

    /**
     * 获取行层级
     *
     * @return 层级
     */
    public Integer getOutlineLevel() {
        return outlineLevel;
    }

    /**
     * 设置行层级
     *
     * @param outlineLevel 层级（不能为负数）
     */
    public void setOutlineLevel(Integer outlineLevel) {
        this.outlineLevel = outlineLevel;
    }

    /**
     * 清除附加属性
     *
     * @return 当前行
     */
    public Row clear() {
        hidden = false;
        outlineLevel = null;
        height = null;
        return this;
    }
}

/*
 * Copyright (c) 2017-2023, guanquan.wang@yandex.com All Rights Reserved.
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

import org.ttzero.excel.drawing.Effect;

import java.nio.file.Path;

/**
 * Picture
 *
 * @author guanquan.wang at 2023-03-08 09:14:39
 */
public class Picture {
    /**
     * Global media id
     */
    public int id;
    /**
     * Relative location name in drawings, not the original name of the file. eq: image1.png
     */
    public String picName;
    /**
     * Axis (one base)
     */
    public int col, row, toCol, toRow;
    /**
     * Padding top | right | bottom | left
     */
    public short[] padding;
    /**
     * Record image position, internal parameters, please do not use
     */
    public int idx;

    // ================Size & Properties================
    /**
     * Size
     */
    public int size;
    /**
     * 0: Move and size with cells
     * 1: Move but don't size with cells
     * 2: Don't move or size with cells
     */
    public int property;
    /**
     * Revolve -360 ~ 360
     */
    public int revolve;

    // ================ Picture Effects ================

    public Effect effect;

    // ================ Picture Local Path ================
    /**
     * 图片临时路径
     */
    public Path localPath;

    /**
     * Padding
     *
     * @param padding int
     * @return current {@code Picture}
     */
    public Picture setPadding(int padding) {
        this.padding = new short[] { (short) padding, (short) padding, (short) padding, (short) padding };
        return this;
    }

    /**
     * Padding Top
     *
     * @param paddingTop int
     * @return current {@code Picture}
     */
    public Picture setPaddingTop(int paddingTop) {
        if (padding == null) padding = new short[] { (short) paddingTop, 0, 0, 0 };
        else padding[0] = (short) paddingTop;
        return this;
    }

    /**
     * Padding Right
     *
     * @param paddingRight int
     * @return current {@code Picture}
     */
    public Picture setPaddingRight(int paddingRight) {
        if (padding == null) padding = new short[] { 0, (short) paddingRight, 0, 0 };
        else padding[1] = (short) paddingRight;
        return this;
    }

    /**
     * Padding Bottom
     *
     * @param paddingBottom int
     * @return current {@code Picture}
     */
    public Picture setPaddingBottom(int paddingBottom) {
        if (padding == null) padding = new short[] { 0, 0, (short) paddingBottom, 0 };
        else padding[2] = (short) paddingBottom;
        return this;
    }

    /**
     * Padding Left
     *
     * @param paddingLeft int
     * @return current {@code Picture}
     */
    public Picture setPaddingLeft(int paddingLeft) {
        if (padding == null) padding = new short[] { 0, 0, 0, (short) paddingLeft };
        else padding[3] = (short) paddingLeft;
        return this;
    }
}

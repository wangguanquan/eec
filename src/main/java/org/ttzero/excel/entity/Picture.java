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

import org.ttzero.excel.drawing.PictureEffect;

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
     * Axis
     */
    public int col, row;
    /**
     * Padding top | right | bottom | left
     */
    public int padding;
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

    public PictureEffect effect;


    /**
     * Padding
     *
     * @param padding int (0 - 255)
     */
    public Picture setPadding(int padding) {
        padding = padding & 0xFF;
        this.padding = padding << 24 | padding << 16 | padding << 8 | padding;
        return this;
    }

    /**
     * Padding Top
     *
     * @param paddingTop int (0 - 255)
     */
    public Picture setPaddingTop(int paddingTop) {
        this.padding = (paddingTop & 0xFF) << 24;
        return this;
    }

    /**
     * Padding Right
     *
     * @param paddingRight int (0 - 255)
     */
    public Picture setPaddingRight(int paddingRight) {
        this.padding = (paddingRight & 0xFF) << 16;
        return this;
    }

    /**
     * Padding Bottom
     *
     * @param paddingBottom int (0 - 255)
     */
    public Picture setPaddingBottom(int paddingBottom) {
        this.padding = (paddingBottom & 0xFF) << 8;
        return this;
    }

    /**
     * Padding Left
     *
     * @param paddingLeft int (0 - 255)
     */
    public Picture setPaddingLeft(int paddingLeft) {
        this.padding = paddingLeft & 0xFF;
        return this;
    }
}

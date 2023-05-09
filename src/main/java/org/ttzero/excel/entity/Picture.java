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

/**
 * Picture
 *
 * @author wangguanquan3 at 2023-03-08 09:14:39
 */
public class Picture {
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
     * -360 ~ 360
     */
    public int revolve;

}

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


package org.ttzero.excel.drawing;

import static org.ttzero.excel.drawing.Enums.PresetBevel;

/**
 * 此元素定义与应用于表中单元格的3D效果关联的棱台的属性。
 *
 * @see Shape3D
 * @author guanquan.wang at 2023-07-25 09:24
 */
public class Bevel {
    /**
     * height: 指定棱台的高度，或者它所应用的形状上方有多远。
     * width: 指定棱台的宽度，或它所应用的形状的距离。
     */
    public double width, height;
    /**
     * 预设棱台
     */
    public PresetBevel prst;

    public double getWidth() {
        return width;
    }

    public Bevel setWidth(double width) {
        this.width = width;
        return this;
    }

    public double getHeight() {
        return height;
    }

    public Bevel setHeight(double height) {
        this.height = height;
        return this;
    }

    public PresetBevel getPrst() {
        return prst;
    }

    public Bevel setPrst(PresetBevel prst) {
        this.prst = prst;
        return this;
    }
}

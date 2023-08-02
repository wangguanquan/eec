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

/**
 * @author guanquan.wang at 2023-07-25 09:24
 */
public class Bevel {
    /**
     * Specifies the bevel width and height
     */
    public double width, height;
    /**
     * Preset Bevel
     */
    public BevelPresetType prst;

    public enum BevelPresetType {
        angle,
        artDeco,
        circle,
        convex,
        coolSlant,
        cross,
        divot,
        hardEdge,
        relaxedInset,
        riblet,
        slope,
        softRound,
    }

    public double getWidth() {
        return width;
    }

    public void setWidth(double width) {
        this.width = width;
    }

    public double getHeight() {
        return height;
    }

    public void setHeight(double height) {
        this.height = height;
    }

    public BevelPresetType getPrst() {
        return prst;
    }

    public void setPrst(BevelPresetType prst) {
        this.prst = prst;
    }
}

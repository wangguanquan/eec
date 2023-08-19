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

import java.awt.Color;

/**
 * @author guanquan.wang at 2023-08-01 19:44
 */
public class Glow {
    /**
     * Specifies shadow color
     */
    public Color color;
    /**
     * Specifies shadow transparency (0-100)
     */
    public int alpha;
    /**
     * Specifies how far to offset the glow (0-150)
     */
    public double dist;

    public Color getColor() {
        return color;
    }

    public Glow setColor(Color color) {
        this.color = color;
        return this;
    }

    public int getAlpha() {
        return alpha;
    }

    public Glow setAlpha(int alpha) {
        this.alpha = alpha;
        return this;
    }

    public double getDist() {
        return dist;
    }

    public Glow setDist(double dist) {
        this.dist = dist;
        return this;
    }
}

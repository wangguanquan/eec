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
 * Creates a partial image reflection near your picture,
 * with a few options for angles and degree of the reflection.
 *
 * @author guanquan.wang at 2023-07-25 09:24
 */
public class Reflection {
    /**
     * Specifies shadow transparency (0-100)
     */
    public int alpha;
    /**
     * Specifies shadow size (0-100)
     */
    public int size;
    /**
     * Specifies the radius of the blur (0-100)
     */
    public double blur;
    /**
     * Specifies how far to offset the shadow (0-100)
     */
    public double dist;
    /**
     * Specifies the direction to offset the shadow (0-360)
     */
    public int direction = 90;

    public int getAlpha() {
        return alpha;
    }

    public Reflection setAlpha(int alpha) {
        this.alpha = alpha;
        return this;
    }

    public int getSize() {
        return size;
    }

    public Reflection setSize(int size) {
        this.size = size;
        return this;
    }

    public double getBlur() {
        return blur;
    }

    public Reflection setBlur(double blur) {
        this.blur = blur;
        return this;
    }

    public double getDist() {
        return dist;
    }

    public Reflection setDist(double dist) {
        this.dist = dist;
        return this;
    }

    public int getDirection() {
        return direction;
    }

    public Reflection setDirection(int direction) {
        this.direction = direction;
        return this;
    }
}

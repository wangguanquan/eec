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

import static org.ttzero.excel.drawing.Enums.Angle;

/**
 * Adds shading at various possible angles to create an effect of depth.
 *
 * @author guanquan.wang at 2023-07-25 09:24
 */
public class Shadow {
    /**
     * Specifies shadow color
     */
    public Color color;
    /**
     * Specifies shadow transparency (0-100)
     */
    public int alpha;
    /**
     * Specifies shadow size (1-200)
     *
     * The effect of {@code size} is equivalent to {@code sx+sy},
     * for example, {@code size=100} is equivalent to {@code sx=100} and {@code sy=100}
     *
     * NOTE: {@code sx} and {@code sy} have higher priority than {@code size}
     */
    public double size = 100; // Default 100%
    /**
     * Specifies the radius of the blur (0-100)
     */
    public double blur;
    /**
     * Specifies shadow alignment. Possible values are b (bottom), bl (bottom left)
     * , br (bottom right), ctr (center), l (left), r (right), t (top), tl (top left)
     * , and tr (top right)
     */
    public Angle angle;
    /**
     * Specifies the direction to offset the shadow (0-360)
     */
    public int direction;
    /**
     * Specifies how far to offset the shadow (0-200)
     */
    public double dist;
    /**
     * Specifies the horizontal skew angle (0-360)
     */
    public double kx;
    /**
     * Specifies the vertical skew angle (0-360)
     */
    public double ky;
    /**
     * Specifies whether the shadow rotates with the shape if the shape is rotated
     */
    public int rotWithShape;
    /**
     * Specifies the horizontal scaling factor (as a percentage). Negative scaling causes a flip. (0 -100)
     */
    public double sx;
    /**
     * Specifies the vertical scaling factor (as a percentage). Negative scaling causes a flip. (0 -100)
     */
    public double sy;

    public Color getColor() {
        return color;
    }

    public Shadow setColor(Color color) {
        this.color = color;
        return this;
    }

    public int getAlpha() {
        return alpha;
    }

    public Shadow setAlpha(int alpha) {
        this.alpha = alpha;
        return this;
    }

    public double getSize() {
        return size;
    }

    public Shadow setSize(double size) {
        this.size = size;
        return this;
    }

    public double getBlur() {
        return blur;
    }

    public Shadow setBlur(double blur) {
        this.blur = blur;
        return this;
    }

    public Angle getAngle() {
        return angle;
    }

    public Shadow setAngle(Angle angle) {
        this.angle = angle;
        return this;
    }

    public int getDirection() {
        return direction;
    }

    public Shadow setDirection(int direction) {
        this.direction = direction;
        return this;
    }

    public double getDist() {
        return dist;
    }

    public Shadow setDist(double dist) {
        this.dist = dist;
        return this;
    }

    public double getKx() {
        return kx;
    }

    public Shadow setKx(double kx) {
        this.kx = kx;
        return this;
    }

    public double getKy() {
        return ky;
    }

    public Shadow setKy(double ky) {
        this.ky = ky;
        return this;
    }

    public int getRotWithShape() {
        return rotWithShape;
    }

    public Shadow setRotWithShape(int rotWithShape) {
        this.rotWithShape = rotWithShape;
        return this;
    }

    public double getSx() {
        return sx;
    }

    public Shadow setSx(double sx) {
        this.sx = sx;
        return this;
    }

    public double getSy() {
        return sy;
    }

    public Shadow setSy(double sy) {
        this.sy = sy;
        return this;
    }
}

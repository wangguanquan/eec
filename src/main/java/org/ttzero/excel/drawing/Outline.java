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

import static org.ttzero.excel.drawing.Enums.Cap;
import static org.ttzero.excel.drawing.Enums.CompoundType;
import static org.ttzero.excel.drawing.Enums.DashPattern;
import static org.ttzero.excel.drawing.Enums.JoinType;

/**
 * The style of the shape's outline is specified with the &lt;a:ln&gt; element.
 * The properties that are determined by this element include such things
 * as the size or weight of the outline, the color, the fill type, and the connector ends.
 * (The style of joints where lines connect is also determined by this element but is not covered here.)
 *
 * @author guanquan.wang at 2023-07-25 14:45
 */
public class Outline {
    public Color color;
    /**
     * Specifies line transparency (0-100)
     */
    public int alpha;
    /**
     * Specifies line width
     */
    public double width;
    public Cap cap;
    public CompoundType cmpd;
    public DashPattern dash;
    public JoinType joinType;
    /**
     * If the {@code JoinType} is set to the {@code Miter}, the MiterLimit property is
     * multiplied by half the {@code width} value to specify a distance at which the
     * intersection of lines is clipped.
     */
    public double miterLimit;

    public Color getColor() {
        return color;
    }

    public Outline setColor(Color color) {
        this.color = color;
        return this;
    }

    public int getAlpha() {
        return alpha;
    }

    public Outline setAlpha(int alpha) {
        this.alpha = alpha;
        return this;
    }

    public double getWidth() {
        return width;
    }

    public Outline setWidth(double width) {
        this.width = width;
        return this;
    }

    public Cap getCap() {
        return cap;
    }

    public Outline setCap(Cap cap) {
        this.cap = cap;
        return this;
    }

    public CompoundType getCmpd() {
        return cmpd;
    }

    public Outline setCmpd(CompoundType cmpd) {
        this.cmpd = cmpd;
        return this;
    }

    public DashPattern getDash() {
        return dash;
    }

    public Outline setDash(DashPattern dash) {
        this.dash = dash;
        return this;
    }

    public JoinType getJoinType() {
        return joinType;
    }

    public Outline setJoinType(JoinType joinType) {
        this.joinType = joinType;
        return this;
    }

    public double getMiterLimit() {
        return miterLimit;
    }

    public Outline setMiterLimit(double miterLimit) {
        this.miterLimit = miterLimit;
        return this;
    }

}

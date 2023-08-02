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

    public enum Cap {
        SQUARE("sq"),
        ROUND("rnd"),
        FLAT("flat")
        ;
        public final String shotName;

        Cap(String shotName) {
            this.shotName = shotName;
        }

        public String getShotName() {
            return shotName;
        }
    }

    public enum CompoundType {
        double_lines("dbl"),
        single_line("sng"),
        thickThin("thickThin"),
        thinThick("thinThick"),
        three_lines("tri")
        ;
        public final String shotName;

        CompoundType(String shotName) {
            this.shotName = shotName;
        }

        public String getShotName() {
            return shotName;
        }
    }

    public enum DashPattern {
        dash,
        dashDot,
        dot,
        lgDash, // large dash
        lgDashDot,
        lgDashDotDot,
        solid,
        sysDash, // system dash
        sysDashDot,
        sysDashDotDot,
        sysDot
    }

    public enum JoinType {
        /**
         * Specifies that a corner where two lines intersect is cut off at a 45 degree angle.
         */
        bevel,
        /**
         * Specifies that a corner where two lines intersect is sharp or clipped,
         * depending on the ShapeOutline.MiterLimit value.
         */
        miter,
        /**
         * Specifies that a corner where two lines intersect is rounded.
         */
        round
    }


    public Color getColor() {
        return color;
    }

    public void setColor(Color color) {
        this.color = color;
    }

    public int getAlpha() {
        return alpha;
    }

    public void setAlpha(int alpha) {
        this.alpha = alpha;
    }

    public double getWidth() {
        return width;
    }

    public void setWidth(double width) {
        this.width = width;
    }

    public Cap getCap() {
        return cap;
    }

    public void setCap(Cap cap) {
        this.cap = cap;
    }

    public CompoundType getCmpd() {
        return cmpd;
    }

    public void setCmpd(CompoundType cmpd) {
        this.cmpd = cmpd;
    }

    public DashPattern getDash() {
        return dash;
    }

    public void setDash(DashPattern dash) {
        this.dash = dash;
    }

    public JoinType getJoinType() {
        return joinType;
    }

    public void setJoinType(JoinType joinType) {
        this.joinType = joinType;
    }

    public double getMiterLimit() {
        return miterLimit;
    }

    public void setMiterLimit(double miterLimit) {
        this.miterLimit = miterLimit;
    }

}

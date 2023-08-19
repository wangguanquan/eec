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
 * @author guanquan.wang at 2023-07-25 14:45
 */
public class Fill {

    /**
     * Solid file
     */
    public static class SolidFill extends Fill {
        /**
         * Specifies fill color
         */
        public Color color;
        /**
         * Specifies fill transparency (0-100)
         */
        public int alpha;

        public int shade;

        public Color getColor() {
            return color;
        }

        public SolidFill setColor(Color color) {
            this.color = color;
            return this;
        }

        public int getAlpha() {
            return alpha;
        }

        public SolidFill setAlpha(int alpha) {
            this.alpha = alpha;
            return this;
        }

        public int getShade() {
            return shade;
        }

        public SolidFill setShade(int shade) {
            this.shade = shade;
            return this;
        }
    }

//    /**
//     * Pattern Fill
//     */
//    public static class PatternFill extends Fill {
//        public Color bgColor, fgColor;
//        public PatternValues pattern;
//    }
//
//    /**
//     * Gradient Fill
//     */
//    public static class GradientFill extends Fill {
//
//    }

}

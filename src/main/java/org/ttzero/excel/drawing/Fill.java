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

        public int shade; // TODO unknown property
    }

    /**
     * Pattern Fill
     */
    public static class PatternFill extends Fill {
        public Color bgColor, fgColor;
        public PatternValues pattern;
    }

    /**
     * Gradient Fill
     */
    public static class GradientFill extends Fill {

    }

    public enum PatternValues {
        /**
         * Cross
         */
        cross,
        /**
         * DarkDownwardDiagonal
         */
        dkDnDiag,
        /**
         * DarkHorizontal
         */
        dkHorz,
        /**
         * DarkUpwardDiagonal
         */
        dkUpDiag,
        /**
         * DarkVertical
         */
        dkVert,
        /**
         * DashedDownwardDiagonal
         */
        dashDnDiag,
        /**
         * DashedHorizontal
         */
        dashHorz,
        /**
         * DashedUpwardDiagonal
         */
        dashUpDiag,
        /**
         * DashedVertical
         */
        dashVert,
        /**
         * DiagonalBrick
         */
        diagBrick,
        /**
         * DiagonalCross
         */
        diagCross,
        /**
         * Divot
         */
        divot,
        /**
         * DotGrid
         */
        dotGrid,
        /**
         * DottedDiamond
         */
        dotDmnd,
        /**
         * DownwardDiagonal
         */
        dnDiag,
        /**
         * Horizontal
         */
        horz,
        /**
         * HorizontalBrick
         */
        horzBrick,
        /**
         * LargeCheck
         */
        lgCheck,
        /**
         * LargeConfetti
         */
        lgConfetti,
        /**
         * LargeGrid
         */
        lgGrid,
        /**
         * LightDownwardDiagonal
         */
        ltDnDiag,
        /**
         * LightHorizontal
         */
        ltHorz,
        /**
         * LightUpwardDiagonal
         */
        ltUpDiag,
        /**
         * LightVertical
         */
        ltVert,
        /**
         * NarrowHorizontal
         */
        narHorz,
        /**
         * NarrowVertical
         */
        narVert,
        /**
         * OpenDiamond
         */
        openDmnd,
        /**
         * Percent10
         */
        pct10,
        /**
         * Percent20
         */
        pct20,
        /**
         * Percent25
         */
        pct25,
        /**
         * Percent30
         */
        pct30,
        /**
         * Percent40
         */
        pct40,
        /**
         * Percent5
         */
        pct5,
        /**
         * Percent50
         */
        pct50,
        /**
         * Percent60
         */
        pct60,
        /**
         * Percent70
         */
        pct70,
        /**
         * Percent75
         */
        pct75,
        /**
         * Percent80
         */
        pct80,
        /**
         * Percent90
         */
        pct90,
        /**
         * Plaid
         */
        plaid,
        /**
         * Shingle
         */
        shingle,
        /**
         * SmallCheck
         */
        smCheck,
        /**
         * SmallConfetti
         */
        smConfetti,
        /**
         * SmallGrid
         */
        smGrid,
        /**
         * SolidDiamond
         */
        solidDmnd,
        /**
         * Sphere
         */
        sphere,
        /**
         * Trellis
         */
        trellis,
        /**
         * UpwardDiagonal
         */
        upDiag,
        /**
         * Vertical
         */
        vert,
        /**
         * Wave
         */
        wave,
        /**
         * Weave
         */
        weave,
        /**
         * WideDownwardDiagonal
         */
        wdDnDiag,
        /**
         * WideUpwardDiagonal
         */
        wdUpDiag,
        /**
         * ZigZag
         */
        zigZag
    }
}

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
 * @author guanquan.wang at 2023-08-17 14:25
 */
public interface Enums {
    enum Angle {
        /**
         * top
         */
        TOP("t"),
        /**
         * left
         */
        LEFT("l"),
        /**
         * bottom
         */
        BOTTOM("b"),
        /**
         * right
         */
        RIGHT("r"),
        /**
         * center
         */
        CENTER("ctr"),
        /**
         * top left
         */
        TOP_LEFT("tl"),
        /**
         * top right
         */
        TOP_RIGHT("tr"),
        /**
         * bottom left
         */
        BOTTOM_LEFT("bl"),
        /**
         * bottom right
         */
        BOTTOM_RIGHT("br")
        ;

        public final String shotName;

        Angle(String shotName) {
            this.shotName = shotName;
        }

        public String getShotName() {
            return shotName;
        }
    }

    enum PresetBevel {
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

    enum PresetCamera {
        isometricBottomDown,
        isometricBottomUp,
        isometricLeftDown,
        isometricLeftUp,
        isometricOffAxis1Left,
        isometricOffAxis1Right,
        isometricOffAxis1Top,
        isometricOffAxis2Left,
        isometricOffAxis2Right,
        isometricOffAxis2Top,
        isometricOffAxis3Bottom,
        isometricOffAxis3Left,
        isometricOffAxis3Right,
        isometricOffAxis4Bottom,
        isometricOffAxis4Left,
        isometricOffAxis4Right,
        isometricRightDown,
        isometricRightUp,
        isometricTopDown,
        isometricTopUp,
        obliqueBottom,
        obliqueBottomLeft,
        obliqueBottomRight,
        obliqueLeft,
        obliqueRight,
        obliqueTop,
        obliqueTopLeft,
        obliqueTopRight,
        orthographicFront,
        perspectiveAbove,
        perspectiveAboveLeftFacing,
        perspectiveAboveRightFacing,
        perspectiveBelow,
        perspectiveContrastingLeftFacing,
        perspectiveContrastingRightFacing,
        perspectiveFront,
        perspectiveHeroicExtremeLeftFacing,
        perspectiveHeroicExtremeRightFacing,
        perspectiveHeroicRightFacing,
        perspectiveLeft,
        perspectiveRelaxed,
        perspectiveRelaxedModerately,
        perspectiveRight,
    }

//    enum PatternValues {
//        /**
//         * Cross
//         */
//        cross,
//        /**
//         * DarkDownwardDiagonal
//         */
//        dkDnDiag,
//        /**
//         * DarkHorizontal
//         */
//        dkHorz,
//        /**
//         * DarkUpwardDiagonal
//         */
//        dkUpDiag,
//        /**
//         * DarkVertical
//         */
//        dkVert,
//        /**
//         * DashedDownwardDiagonal
//         */
//        dashDnDiag,
//        /**
//         * DashedHorizontal
//         */
//        dashHorz,
//        /**
//         * DashedUpwardDiagonal
//         */
//        dashUpDiag,
//        /**
//         * DashedVertical
//         */
//        dashVert,
//        /**
//         * DiagonalBrick
//         */
//        diagBrick,
//        /**
//         * DiagonalCross
//         */
//        diagCross,
//        /**
//         * Divot
//         */
//        divot,
//        /**
//         * DotGrid
//         */
//        dotGrid,
//        /**
//         * DottedDiamond
//         */
//        dotDmnd,
//        /**
//         * DownwardDiagonal
//         */
//        dnDiag,
//        /**
//         * Horizontal
//         */
//        horz,
//        /**
//         * HorizontalBrick
//         */
//        horzBrick,
//        /**
//         * LargeCheck
//         */
//        lgCheck,
//        /**
//         * LargeConfetti
//         */
//        lgConfetti,
//        /**
//         * LargeGrid
//         */
//        lgGrid,
//        /**
//         * LightDownwardDiagonal
//         */
//        ltDnDiag,
//        /**
//         * LightHorizontal
//         */
//        ltHorz,
//        /**
//         * LightUpwardDiagonal
//         */
//        ltUpDiag,
//        /**
//         * LightVertical
//         */
//        ltVert,
//        /**
//         * NarrowHorizontal
//         */
//        narHorz,
//        /**
//         * NarrowVertical
//         */
//        narVert,
//        /**
//         * OpenDiamond
//         */
//        openDmnd,
//        /**
//         * Percent10
//         */
//        pct10,
//        /**
//         * Percent20
//         */
//        pct20,
//        /**
//         * Percent25
//         */
//        pct25,
//        /**
//         * Percent30
//         */
//        pct30,
//        /**
//         * Percent40
//         */
//        pct40,
//        /**
//         * Percent5
//         */
//        pct5,
//        /**
//         * Percent50
//         */
//        pct50,
//        /**
//         * Percent60
//         */
//        pct60,
//        /**
//         * Percent70
//         */
//        pct70,
//        /**
//         * Percent75
//         */
//        pct75,
//        /**
//         * Percent80
//         */
//        pct80,
//        /**
//         * Percent90
//         */
//        pct90,
//        /**
//         * Plaid
//         */
//        plaid,
//        /**
//         * Shingle
//         */
//        shingle,
//        /**
//         * SmallCheck
//         */
//        smCheck,
//        /**
//         * SmallConfetti
//         */
//        smConfetti,
//        /**
//         * SmallGrid
//         */
//        smGrid,
//        /**
//         * SolidDiamond
//         */
//        solidDmnd,
//        /**
//         * Sphere
//         */
//        sphere,
//        /**
//         * Trellis
//         */
//        trellis,
//        /**
//         * UpwardDiagonal
//         */
//        upDiag,
//        /**
//         * Vertical
//         */
//        vert,
//        /**
//         * Wave
//         */
//        wave,
//        /**
//         * Weave
//         */
//        weave,
//        /**
//         * WideDownwardDiagonal
//         */
//        wdDnDiag,
//        /**
//         * WideUpwardDiagonal
//         */
//        wdUpDiag,
//        /**
//         * ZigZag
//         */
//        zigZag
//    }

    enum Rig {
        balanced,
        brightRoom,
        chilly,
        contrasting,
        flat,
        flood,
        freezing,
        glow,
        harsh,
        legacyFlat1,
        legacyFlat2,
        legacyFlat3,
        legacyFlat4,
        legacyHarsh1,
        legacyHarsh2,
        legacyHarsh3,
        legacyHarsh4,
        legacyNormal1,
        legacyNormal2,
        legacyNormal3,
        legacyNormal4,
        morning,
        soft,
        sunrise,
        sunset,
        threePt,
        twoPt
    }

    enum Cap {
        square("sq"),
        round("rnd"),
        flat("flat")
        ;
        public final String shotName;

        Cap(String shotName) {
            this.shotName = shotName;
        }

        public String getShotName() {
            return shotName;
        }
    }

    enum CompoundType {
        doubleLines("dbl"),
        singleLine("sng"),
        thickThin("thickThin"),
        thinThick("thinThick"),
        threeLines("tri")
        ;
        public final String shotName;

        CompoundType(String shotName) {
            this.shotName = shotName;
        }

        public String getShotName() {
            return shotName;
        }
    }

    enum DashPattern {
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

    enum JoinType {
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

    enum Material {
        clear,
        dkEdge,
        flat,
        legacyMatte,
        legacyMetal,
        legacyPlastic,
        legacyWireframe,
        matte,
        metal,
        plastic,
        powder,
        softEdge,
        softMetal,
        translucentPowder,
        warmMatte,
    }

    enum ShapeType {
        /**
         * AccentBorderCallout1
         */
        accentBorderCallout1,
        /**
         * AccentBorderCallout2
         */
        accentBorderCallout2,
        /**
         * AccentBorderCallout3
         */
        accentBorderCallout3,
        /**
         * AccentCallout1
         */
        accentCallout1,
        /**
         * AccentCallout2
         */
        accentCallout2,
        /**
         * AccentCallout3
         */
        accentCallout3,
        /**
         * ActionButtonBackPrevious
         */
        actionButtonBackPrevious,
        /**
         * ActionButtonBeginning
         */
        actionButtonBeginning,
        /**
         * ActionButtonBlank
         */
        actionButtonBlank,
        /**
         * ActionButtonDocument
         */
        actionButtonDocument,
        /**
         * ActionButtonEnd
         */
        actionButtonEnd,
        /**
         * ActionButtonForwardNext
         */
        actionButtonForwardNext,
        /**
         * ActionButtonHelp
         */
        actionButtonHelp,
        /**
         * ActionButtonHome
         */
        actionButtonHome,
        /**
         * ActionButtonInformation
         */
        actionButtonInformation,
        /**
         * ActionButtonMovie
         */
        actionButtonMovie,
        /**
         * ActionButtonReturn
         */
        actionButtonReturn,
        /**
         * ActionButtonSound
         */
        actionButtonSound,
        /**
         * Arc
         */
        arc,
        /**
         * BentArrow
         */
        bentArrow,
        /**
         * BentConnector2
         */
        bentConnector2,
        /**
         * BentConnector3
         */
        bentConnector3,
        /**
         * BentConnector4
         */
        bentConnector4,
        /**
         * BentConnector5
         */
        bentConnector5,
        /**
         * BentUpArrow
         */
        bentUpArrow,
        /**
         * Bevel
         */
        bevel,
        /**
         * BlockArc
         */
        blockArc,
        /**
         * BorderCallout1
         */
        borderCallout1,
        /**
         * BorderCallout2
         */
        borderCallout2,
        /**
         * BorderCallout3
         */
        borderCallout3,
        /**
         * BracePair
         */
        bracePair,
        /**
         * BracketPair
         */
        bracketPair,
        /**
         * Callout1
         */
        callout1,
        /**
         * Callout2
         */
        callout2,
        /**
         * Callout3
         */
        callout3,
        /**
         * Can
         */
        can,
        /**
         * ChartPlus
         */
        chartPlus,
        /**
         * ChartStar
         */
        chartStar,
        /**
         * ChartX
         */
        chartX,
        /**
         * Chevron
         */
        chevron,
        /**
         * Chord
         */
        chord,
        /**
         * CircularArrow
         */
        circularArrow,
        /**
         * Cloud
         */
        cloud,
        /**
         * CloudCallout
         */
        cloudCallout,
        /**
         * Corner
         */
        corner,
        /**
         * CornerTabs
         */
        cornerTabs,
        /**
         * Cube
         */
        cube,
        /**
         * CurvedConnector2
         */
        curvedConnector2,
        /**
         * CurvedConnector3
         */
        curvedConnector3,
        /**
         * CurvedConnector4
         */
        curvedConnector4,
        /**
         * CurvedConnector5
         */
        curvedConnector5,
        /**
         * CurvedDownArrow
         */
        curvedDownArrow,
        /**
         * CurvedLeftArrow
         */
        curvedLeftArrow,
        /**
         * CurvedRightArrow
         */
        curvedRightArrow,
        /**
         * CurvedUpArrow
         */
        curvedUpArrow,
        /**
         * Decagon
         */
        decagon,
        /**
         * DiagonalStripe
         */
        diagStripe,
        /**
         * Diamond
         */
        diamond,
        /**
         * Dodecagon
         */
        dodecagon,
        /**
         * Donut
         */
        donut,
        /**
         * DoubleWave
         */
        doubleWave,
        /**
         * DownArrow
         */
        downArrow,
        /**
         * DownArrowCallout
         */
        downArrowCallout,
        /**
         * Ellipse
         */
        ellipse,
        /**
         * EllipseRibbon
         */
        ellipseRibbon,
        /**
         * EllipseRibbon2
         */
        ellipseRibbon2,
        /**
         * FlowChartAlternateProcess
         */
        flowChartAlternateProcess,
        /**
         * FlowChartCollate
         */
        flowChartCollate,
        /**
         * FlowChartConnector
         */
        flowChartConnector,
        /**
         * FlowChartDecision
         */
        flowChartDecision,
        /**
         * FlowChartDelay
         */
        flowChartDelay,
        /**
         * FlowChartDisplay
         */
        flowChartDisplay,
        /**
         * FlowChartDocument
         */
        flowChartDocument,
        /**
         * FlowChartExtract
         */
        flowChartExtract,
        /**
         * FlowChartInputOutput
         */
        flowChartInputOutput,
        /**
         * FlowChartInternalStorage
         */
        flowChartInternalStorage,
        /**
         * FlowChartMagneticDisk
         */
        flowChartMagneticDisk,
        /**
         * FlowChartMagneticDrum
         */
        flowChartMagneticDrum,
        /**
         * FlowChartMagneticTape
         */
        flowChartMagneticTape,
        /**
         * FlowChartManualInput
         */
        flowChartManualInput,
        /**
         * FlowChartManualOperation
         */
        flowChartManualOperation,
        /**
         * FlowChartMerge
         */
        flowChartMerge,
        /**
         * FlowChartMultidocument
         */
        flowChartMultidocument,
        /**
         * FlowChartOfflineStorage
         */
        flowChartOfflineStorage,
        /**
         * FlowChartOffpageConnector
         */
        flowChartOffpageConnector,
        /**
         * FlowChartOnlineStorage
         */
        flowChartOnlineStorage,
        /**
         * FlowChartOr
         */
        flowChartOr,
        /**
         * FlowChartPredefinedProcess
         */
        flowChartPredefinedProcess,
        /**
         * FlowChartPreparation
         */
        flowChartPreparation,
        /**
         * FlowChartProcess
         */
        flowChartProcess,
        /**
         * FlowChartPunchedCard
         */
        flowChartPunchedCard,
        /**
         * FlowChartPunchedTape
         */
        flowChartPunchedTape,
        /**
         * FlowChartSort
         */
        flowChartSort,
        /**
         * FlowChartSummingJunction
         */
        flowChartSummingJunction,
        /**
         * FlowChartTerminator
         */
        flowChartTerminator,
        /**
         * FoldedCorner
         */
        foldedCorner,
        /**
         * Frame
         */
        frame,
        /**
         * Funnel
         */
        funnel,
        /**
         * Gear6
         */
        gear6,
        /**
         * Gear9
         */
        gear9,
        /**
         * HalfFrame
         */
        halfFrame,
        /**
         * Heart
         */
        heart,
        /**
         * Heptagon
         */
        heptagon,
        /**
         * Hexagon
         */
        hexagon,
        /**
         * HomePlate
         */
        homePlate,
        /**
         * HorizontalScroll
         */
        horizontalScroll,
        /**
         * IrregularSeal1
         */
        irregularSeal1,
        /**
         * IrregularSeal2
         */
        irregularSeal2,
        /**
         * LeftArrow
         */
        leftArrow,
        /**
         * LeftArrowCallout
         */
        leftArrowCallout,
        /**
         * LeftBrace
         */
        leftBrace,
        /**
         * LeftBracket
         */
        leftBracket,
        /**
         * LeftCircularArrow
         */
        leftCircularArrow,
        /**
         * LeftRightArrow
         */
        leftRightArrow,
        /**
         * LeftRightArrowCallout
         */
        leftRightArrowCallout,
        /**
         * LeftRightCircularArrow
         */
        leftRightCircularArrow,
        /**
         * LeftRightRibbon
         */
        leftRightRibbon,
        /**
         * LeftRightUpArrow
         */
        leftRightUpArrow,
        /**
         * LeftUpArrow
         */
        leftUpArrow,
        /**
         * LightningBolt
         */
        lightningBolt,
        /**
         * Line
         */
        line,
        /**
         * LineInverse
         */
        lineInv,
        /**
         * MathDivide
         */
        mathDivide,
        /**
         * MathEqual
         */
        mathEqual,
        /**
         * MathMinus
         */
        mathMinus,
        /**
         * MathMultiply
         */
        mathMultiply,
        /**
         * MathNotEqual
         */
        mathNotEqual,
        /**
         * MathPlus
         */
        mathPlus,
        /**
         * Moon
         */
        moon,
        /**
         * NonIsoscelesTrapezoid
         */
        nonIsoscelesTrapezoid,
        /**
         * NoSmoking
         */
        noSmoking,
        /**
         * NotchedRightArrow
         */
        notchedRightArrow,
        /**
         * Octagon
         */
        octagon,
        /**
         * Parallelogram
         */
        parallelogram,
        /**
         * Pentagon
         */
        pentagon,
        /**
         * Pie
         */
        pie,
        /**
         * PieWedge
         */
        pieWedge,
        /**
         * Plaque
         */
        plaque,
        /**
         * PlaqueTabs
         */
        plaqueTabs,
        /**
         * Plus
         */
        plus,
        /**
         * QuadArrow
         */
        quadArrow,
        /**
         * QuadArrowCallout
         */
        quadArrowCallout,
        /**
         * Rectangle
         */
        rect,
        /**
         * Ribbon
         */
        ribbon,
        /**
         * Ribbon2
         */
        ribbon2,
        /**
         * RightArrow
         */
        rightArrow,
        /**
         * RightArrowCallout
         */
        rightArrowCallout,
        /**
         * RightBrace
         */
        rightBrace,
        /**
         * RightBracket
         */
        rightBracket,
        /**
         * RightTriangle
         */
        rtTriangle,
        /**
         * Round1Rectangle
         */
        round1Rect,
        /**
         * Round2DiagonalRectangle
         */
        round2DiagRect,
        /**
         * Round2SameRectangle
         */
        round2SameRect,
        /**
         * RoundRectangle
         */
        roundRect,
        /**
         * SmileyFace
         */
        smileyFace,
        /**
         * Snip1Rectangle
         */
        snip1Rect,
        /**
         * Snip2DiagonalRectangle
         */
        snip2DiagRect,
        /**
         * Snip2SameRectangle
         */
        snip2SameRect,
        /**
         * SnipRoundRectangle
         */
        snipRoundRect,
        /**
         * SquareTabs
         */
        squareTabs,
        /**
         * Star10
         */
        star10,
        /**
         * Star12
         */
        star12,
        /**
         * Star16
         */
        star16,
        /**
         * Star24
         */
        star24,
        /**
         * Star32
         */
        star32,
        /**
         * Star4
         */
        star4,
        /**
         * Star5
         */
        star5,
        /**
         * Star6
         */
        star6,
        /**
         * Star7
         */
        star7,
        /**
         * Star8
         */
        star8,
        /**
         * StraightConnector1
         */
        straightConnector1,
        /**
         * StripedRightArrow
         */
        stripedRightArrow,
        /**
         * Sun
         */
        sun,
        /**
         * SwooshArrow
         */
        swooshArrow,
        /**
         * Teardrop
         */
        teardrop,
        /**
         * Trapezoid
         */
        trapezoid,
        /**
         * Triangle
         */
        triangle,
        /**
         * UpArrow
         */
        upArrow,
        /**
         * UpArrowCallout
         */
        upArrowCallout,
        /**
         * UpDownArrow
         */
        upDownArrow,
        /**
         * UpDownArrowCallout
         */
        upDownArrowCallout,
        /**
         * UTurnArrow
         */
        uturnArrow,
        /**
         * VerticalScroll
         */
        verticalScroll,
        /**
         * Wave
         */
        wave,
        /**
         * WedgeEllipseCallout
         */
        wedgeEllipseCallout,
        /**
         * WedgeRectangleCallout
         */
        wedgeRectCallout,
        /**
         * WedgeRoundRectangleCallout
         */
        wedgeRoundRectCallout,
    }

}

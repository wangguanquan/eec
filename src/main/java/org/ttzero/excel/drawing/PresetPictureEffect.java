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

import org.ttzero.excel.manager.docProps.Tuple2;

import java.awt.Color;
import java.util.Arrays;
import java.util.Collections;


import static org.ttzero.excel.drawing.Enums.Angle;
import static org.ttzero.excel.drawing.Enums.Cap;
import static org.ttzero.excel.drawing.Enums.CompoundType;
import static org.ttzero.excel.drawing.Enums.DashPattern;
import static org.ttzero.excel.drawing.Enums.JoinType;
import static org.ttzero.excel.drawing.Enums.Material;
import static org.ttzero.excel.drawing.Enums.PresetBevel;
import static org.ttzero.excel.drawing.Enums.PresetCamera;
import static org.ttzero.excel.drawing.Enums.Rig;
import static org.ttzero.excel.drawing.Enums.ShapeType;

/**
 * Preset Picture Effects
 *
 * @author guanquan.wang at 2023-07-25 09:59
 */
public enum PresetPictureEffect implements EffectProducer {
    // 0
    None {
        @Override public Effect getEffect() {
            return null;
        }
    },
    // 1
    SimpleFrame_White {
        @Override public Effect getEffect() {
            if (effect != null) return effect;
            effect = new Effect();

            Fill.SolidFill fill = new Fill.SolidFill();
            fill.color = Color.WHITE;
            fill.shade = 85;
            effect.fill = fill;

            Outline ln = new Outline();
            effect.outline = ln;
            ln.width = 7D;
            ln.color = Color.WHITE;
            ln.cap = Cap.square;
            ln.joinType = JoinType.miter;
            ln.miterLimit = 800D;

            Shadow shadow = new Shadow();
            effect.shadow = shadow;
            shadow.color = Color.BLACK;
            shadow.alpha = 60;
            shadow.blur = 4.33D;
            shadow.direction = 90;
            shadow.dist = 1.42D;
            shadow.angle = Angle.TOP_LEFT;

            effect.geometry = ShapeType.rect;

            Scene3D scene = new Scene3D();
            effect.scene3D = scene;
            Camera camera = new Camera();
            camera.presetCamera = PresetCamera.orthographicFront;
            scene.camera = camera;
            LightRig lightRig = new LightRig();
            lightRig.rig = Rig.twoPt;
            lightRig.angle = Angle.TOP;
            lightRig.revolution = 120D;
            scene.lightRig = lightRig;

            Shape3D shape = new Shape3D();
            effect.shape3D = shape;
            Bevel bevel = new Bevel();
            shape.bevelTop = bevel;
            bevel.width = 2D;
            bevel.height = 1.5D;
            shape.contourColor = Color.WHITE;
            return effect;
        }
    },
    // 2
    BeveledMatte_White {
        @Override public Effect getEffect() {
            if (effect != null) return effect;
            effect = new Effect();

            Fill.SolidFill fill = new Fill.SolidFill();
            fill.color = Color.WHITE;
            fill.shade = 85;
            effect.fill = fill;

            Outline ln = new Outline();
            effect.outline = ln;
            ln.width = 15D;
            ln.color = Color.WHITE;
            ln.cap = Cap.round;

            Shadow shadow = new Shadow();
            effect.shadow = shadow;
            shadow.color = Color.BLACK;
            shadow.alpha = 59;
            shadow.blur = 3.94D;
            shadow.angle = Angle.TOP_LEFT;

            effect.geometry = ShapeType.rect;

            Scene3D scene = new Scene3D();
            effect.scene3D = scene;
            Camera camera = new Camera();
            camera.presetCamera = PresetCamera.orthographicFront;
            scene.camera = camera;
            LightRig lightRig = new LightRig();
            lightRig.rig = Rig.twoPt;
            lightRig.angle = Angle.TOP;
            lightRig.revolution = 130D;
            scene.lightRig = lightRig;

            Shape3D shape = new Shape3D();
            effect.shape3D = shape;
            Bevel bevel = new Bevel();
            shape.bevelTop = bevel;
            bevel.width = 4D;
            bevel.height = 1.3D;
            shape.contourColor = new Color(192, 192, 192);
            shape.contourWidth = 0.5D;
            return effect;
        }
    },
    // 3
    MetalFrame {
        @Override public Effect getEffect() {
            if (effect != null) return effect;
            effect = new Effect();

            Outline ln = new Outline();
            effect.outline = ln;
            ln.width = 15D;
            ln.color = new Color(200, 198, 189);
            ln.cap = Cap.square;
            ln.dash = DashPattern.solid;
            ln.joinType = JoinType.miter;
            ln.miterLimit = 800D;

            Shadow shadow = new Shadow();
            effect.shadow = shadow;
            shadow.color = Color.BLACK;
            shadow.alpha = 57;
            shadow.blur = 20D;
            shadow.angle = Angle.BOTTOM_LEFT;

            effect.geometry = ShapeType.rect;

            Scene3D scene = new Scene3D();
            effect.scene3D = scene;
            Camera camera = new Camera();
            camera.presetCamera = PresetCamera.perspectiveFront;
            camera.fov = 90D;
            scene.camera = camera;
            LightRig lightRig = new LightRig();
            lightRig.rig = Rig.threePt;
            lightRig.angle = Angle.TOP;
            lightRig.revolution = 35D;
            scene.lightRig = lightRig;

            Shape3D shape = new Shape3D();
            effect.shape3D = shape;
            Bevel bevel = new Bevel();
            shape.bevelTop = bevel;
            bevel.width = 24D;
            bevel.height = 12D;
            bevel.prst = PresetBevel.hardEdge;
            shape.extrusionColor = Color.BLACK;
            shape.extrusionHeight = 2D;
            return effect;
        }
    },
    // 4
    DropShadowRectangle {
        @Override public Effect getEffect() {
            if (effect != null) return effect;
            effect = new Effect();

            Shadow shadow = new Shadow();
            shadow.color = new Color(51, 51, 51);
            shadow.alpha = 35;
            shadow.direction = 45;
            shadow.blur = 23D;
            shadow.dist = 11D;
            shadow.angle = Angle.TOP_LEFT;
            effect.shadow = shadow;

            effect.geometry = ShapeType.rect;
            return effect;
        }
    },
    // 5
    ReflectedRoundedRectangle {
        @Override public Effect getEffect() {
            if (effect != null) return effect;
            effect = new Effect();

            Fill.SolidFill fill = new Fill.SolidFill();
            fill.color = Color.WHITE;
            fill.shade = 85;
            effect.fill = fill;

            Reflection reflection = new Reflection();
            reflection.blur = 1D;
            reflection.alpha = 62;
            reflection.size = 28D;
            reflection.dist = 0.4D;
            effect.reflection = reflection;

            effect.geometry = ShapeType.roundRect;
            effect.geometryAdjustValueList = Collections.singletonList(Tuple2.of("adj", "val 8594"));
            return effect;
        }
    },
    // 6
    SoftEdgeRectangle {
        @Override public Effect getEffect() {
            if (effect != null) return effect;
            effect = new Effect();

            effect.softEdges = 8.86D;
            effect.geometry = ShapeType.rect;
            return effect;
        }
    },
    // 7
    DoubleFrame_Black {
        @Override public Effect getEffect() {
            if (effect != null) return effect;
            effect = new Effect();

            Outline ln = new Outline();
            effect.outline = ln;
            ln.width = 18D;
            ln.color = Color.BLACK;
            ln.cap = Cap.square;
            ln.cmpd = CompoundType.thickThin;
            ln.dash = DashPattern.solid;
            ln.joinType = JoinType.miter;
            ln.miterLimit = 800D;

            Shadow shadow = new Shadow();
            effect.innerShadow = shadow;
            shadow.color = Color.BLACK;
            shadow.blur = 6D;

            effect.geometry = ShapeType.rect;
            return effect;
        }
    },
    // 8
    ThickMatte_Black {
        @Override public Effect getEffect() {
            if (effect != null) return effect;
            effect = new Effect();

            Fill.SolidFill fill = new Fill.SolidFill();
            fill.color = Color.BLACK;
            fill.shade = 95;
            effect.fill = fill;

            Outline ln = new Outline();
            effect.outline = ln;
            ln.width = 35D;
            ln.color = Color.BLACK;
            ln.cap = Cap.square;
            ln.joinType = JoinType.miter;
            ln.miterLimit = 800D;

            Shadow shadow = new Shadow();
            effect.shadow = shadow;
            shadow.color = Color.BLACK;
            shadow.alpha = 60;
            shadow.blur = 20D;
            shadow.direction = 45;
            shadow.dist = 15D;
            shadow.angle = Angle.BOTTOM_LEFT;
            shadow.sy = 90D;

            effect.geometry = ShapeType.rect;
            return effect;
        }
    },
    // 9
    SimpleFrame_Black {
        @Override public Effect getEffect() {
            if (effect != null) return effect;
            effect = new Effect();

            Outline ln = new Outline();
            effect.outline = ln;
            ln.width = 3D;
            ln.color = Color.BLACK;
            ln.cap = Cap.square;
            ln.dash = DashPattern.solid;
            ln.joinType = JoinType.miter;
            ln.miterLimit = 800D;

            Shadow shadow = new Shadow();
            effect.shadow = shadow;
            shadow.color = Color.BLACK;
            shadow.alpha = 57;
            shadow.blur = 4D;
            shadow.direction = 45;
            shadow.dist = 3D;
            shadow.angle = Angle.TOP_LEFT;

            effect.geometry = ShapeType.rect;
            return effect;
        }
    },
    // 10
    BeveledOval_Black {
        @Override public Effect getEffect() {
            if (effect != null) return effect;
            effect = new Effect();

            Outline ln = new Outline();
            effect.outline = ln;
            ln.width = 5D;
            ln.color = new Color(51, 51, 51);
            ln.cap = Cap.round;

            Shadow shadow = new Shadow();
            effect.shadow = shadow;
            shadow.color = Color.BLACK;
            shadow.alpha = 78;
            shadow.blur = 30D;
            shadow.direction = 90;
            shadow.dist = 23D;
            shadow.sx = -80D;
            shadow.sy = -18D;

            effect.geometry = ShapeType.ellipse;

            Scene3D scene = new Scene3D();
            effect.scene3D = scene;
            Camera camera = new Camera();
            camera.presetCamera = PresetCamera.orthographicFront;
            scene.camera = camera;
            LightRig lightRig = new LightRig();
            lightRig.rig = Rig.contrasting;
            lightRig.angle = Angle.TOP;
            lightRig.revolution = 50D;
            scene.lightRig = lightRig;

            Shape3D shape = new Shape3D();
            effect.shape3D = shape;
            Bevel bevel = new Bevel();
            shape.bevelTop = bevel;
            bevel.width = 7.5D;
            bevel.height = 2.5D;
            shape.contourColor = new Color(51, 51, 51);
            shape.contourWidth = 0.6D;
            return effect;
        }
    },
    // 11
    CompoundFrame_Black {
        @Override public Effect getEffect() {
            if (effect != null) return effect;
            effect = new Effect();

            Outline ln = new Outline();
            effect.outline = ln;
            ln.width = 7D;
            ln.color = Color.BLACK;
            ln.cap = Cap.square;
            ln.dash = DashPattern.solid;
            ln.cmpd = CompoundType.thickThin;
            ln.joinType = JoinType.miter;
            ln.miterLimit = 800D;

            Shadow shadow = new Shadow();
            effect.innerShadow = shadow;
            shadow.color = Color.BLACK;
            shadow.blur = 6D;

            effect.geometry = ShapeType.rect;
            return effect;
        }
    },
    // 12
    ModerateFrame_Black {
        @Override public Effect getEffect() {
            if (effect != null) return effect;
            effect = new Effect();

            Outline ln = new Outline();
            effect.outline = ln;
            ln.width = 10D;
            ln.color = Color.BLACK;
            ln.cap = Cap.square;
            ln.joinType = JoinType.miter;
            ln.miterLimit = 800D;

            Shadow shadow = new Shadow();
            effect.shadow = shadow;
            shadow.color = Color.BLACK;
            shadow.alpha = 60;
            shadow.blur = 4.5D;
            shadow.direction = 45;
            shadow.dist = 4D;
            shadow.angle = Angle.TOP_LEFT;

            effect.geometry = ShapeType.rect;
            return effect;
        }
    },
    // 13
    CenterShadowRectangle {
        @Override public Effect getEffect() {
            if (effect != null) return effect;
            effect = new Effect();

            Shadow shadow = new Shadow();
            shadow.color = Color.BLACK;
            shadow.alpha = 30;
            shadow.blur = 15D;
            shadow.angle = Angle.TOP_LEFT;
            effect.shadow = shadow;

            effect.geometry = ShapeType.rect;
            return effect;
        }
    },
    // 14
    RoundedDiagonalCorner_White {
        @Override public Effect getEffect() {
            if (effect != null) return effect;
            effect = new Effect();

            Shadow shadow = new Shadow();
            shadow.color = Color.BLACK;
            shadow.alpha = 57;
            shadow.blur = 20D;
            shadow.angle = Angle.TOP_LEFT;
            effect.shadow = shadow;

            effect.geometry = ShapeType.round2DiagRect;
            effect.geometryAdjustValueList = Arrays.asList(Tuple2.of("adj1", "val 16667"), Tuple2.of("adj2", "val 0"));

            Outline ln = new Outline();
            effect.outline = ln;
            ln.width = 7D;
            ln.color = Color.WHITE;
            ln.cap = Cap.square;
            ln.joinType = JoinType.miter;
            ln.miterLimit = 800D;
            return effect;
        }
    },
    // 15
    SnipDiagonalCorner_White {
        @Override public Effect getEffect() {
            if (effect != null) return effect;
            effect = new Effect();

            Fill.SolidFill fill = new Fill.SolidFill();
            fill.color = Color.WHITE;
            fill.shade = 85;
            effect.fill = fill;

            Shadow shadow = new Shadow();
            shadow.color = Color.BLACK;
            shadow.alpha = 55;
            shadow.blur = 7D;
            shadow.angle = Angle.TOP_LEFT;
            effect.shadow = shadow;

            effect.geometry = ShapeType.snip2DiagRect;

            Outline ln = new Outline();
            effect.outline = ln;
            ln.width = 7D;
            ln.color = Color.WHITE;
            ln.cap = Cap.square;
            ln.joinType = JoinType.miter;
            ln.miterLimit = 800D;

            Scene3D scene = new Scene3D();
            effect.scene3D = scene;
            Camera camera = new Camera();
            camera.presetCamera = PresetCamera.orthographicFront;
            scene.camera = camera;
            LightRig lightRig = new LightRig();
            lightRig.rig = Rig.twoPt;
            lightRig.angle = Angle.TOP;
            lightRig.revolution = 120D;
            scene.lightRig = lightRig;

            Shape3D shape = new Shape3D();
            effect.shape3D = shape;
            Bevel bevel = new Bevel();
            shape.bevelTop = bevel;
            bevel.width = 2D;
            bevel.height = 1.5D;
            shape.contourColor = Color.WHITE;
            return effect;
        }
    },
    // 16
    ModerateFrame_White {
        @Override public Effect getEffect() {
            if (effect != null) return effect;
            effect = new Effect();

            Fill.SolidFill fill = new Fill.SolidFill();
            fill.color = Color.WHITE;
            fill.shade = 85;
            effect.fill = fill;

            Shadow shadow = new Shadow();
            shadow.color = Color.BLACK;
            shadow.alpha = 60;
            shadow.blur = 4.33D;
            shadow.direction = 90;
            shadow.dist = 1.4D;
            shadow.angle = Angle.TOP_LEFT;
            effect.shadow = shadow;

            effect.geometry = ShapeType.rect;

            Outline ln = new Outline();
            effect.outline = ln;
            ln.width = 15D;
            ln.color = Color.WHITE;
            ln.cap = Cap.square;
            ln.joinType = JoinType.miter;
            ln.miterLimit = 800D;

            Scene3D scene = new Scene3D();
            effect.scene3D = scene;
            Camera camera = new Camera();
            camera.presetCamera = PresetCamera.orthographicFront;
            scene.camera = camera;
            LightRig lightRig = new LightRig();
            lightRig.rig = Rig.twoPt;
            lightRig.angle = Angle.TOP;
            lightRig.revolution = 120D;
            scene.lightRig = lightRig;

            Shape3D shape = new Shape3D();
            effect.shape3D = shape;
            Bevel bevel = new Bevel();
            shape.bevelTop = bevel;
            bevel.width = 2D;
            bevel.height = 1.5D;
            shape.contourColor = Color.WHITE;
            return effect;
        }
    },
    // 17
    Rotated_White {
        @Override public Effect getEffect() {
            if (effect != null) return effect;
            effect = new Effect();

            Fill.SolidFill fill = new Fill.SolidFill();
            fill.color = Color.WHITE;
            fill.shade = 85;
            effect.fill = fill;

            Shadow shadow = new Shadow();
            shadow.color = Color.BLACK;
            shadow.alpha = 70;
            shadow.blur = 5.12D;
            shadow.direction = 215;
            shadow.dist = 4D;
            shadow.angle = Angle.TOP_LEFT;
            shadow.kx = 3.25D;
            shadow.ky = 2.42D;
            effect.shadow = shadow;

            effect.geometry = ShapeType.rect;

            Outline ln = new Outline();
            effect.outline = ln;
            ln.width = 15D;
            ln.color = Color.WHITE;
            ln.cap = Cap.square;
            ln.joinType = JoinType.miter;
            ln.miterLimit = 800D;

            Scene3D scene = new Scene3D();
            effect.scene3D = scene;
            Camera camera = new Camera();
            camera.presetCamera = PresetCamera.orthographicFront;
            camera.revolution = 6D;
            scene.camera = camera;
            LightRig lightRig = new LightRig();
            lightRig.rig = Rig.twoPt;
            lightRig.angle = Angle.TOP;
            lightRig.revolution = 120D;
            scene.lightRig = lightRig;

            Shape3D shape = new Shape3D();
            effect.shape3D = shape;
            Bevel bevel = new Bevel();
            shape.bevelTop = bevel;
            bevel.width = 2D;
            bevel.height = 1.5D;
            shape.contourColor = new Color(150, 150, 150);
            shape.contourWidth = 1D;
            return effect;
        }
    },
    // 18
    PerspectiveShadow_White {
        @Override public Effect getEffect() {
            if (effect != null) return effect;
            effect = new Effect();

            Shadow shadow = new Shadow();
            shadow.color = Color.BLACK;
            shadow.alpha = 80;
            shadow.blur = 6D;
            shadow.direction = 175;
            shadow.dist = 7.5D;
            shadow.angle = Angle.BOTTOM_RIGHT;
            shadow.kx = 15D;
            shadow.sx = 97D;
            shadow.sy = 23D;
            effect.shadow = shadow;

            effect.geometry = ShapeType.rect;

            Outline ln = new Outline();
            effect.outline = ln;
            ln.width = 10D;
            ln.color = Color.WHITE;
            ln.cap = Cap.round;

            Scene3D scene = new Scene3D();
            effect.scene3D = scene;
            Camera camera = new Camera();
            camera.presetCamera = PresetCamera.orthographicFront;
            scene.camera = camera;
            LightRig lightRig = new LightRig();
            lightRig.rig = Rig.twoPt;
            lightRig.angle = Angle.TOP;
            lightRig.revolution = 130D;
            scene.lightRig = lightRig;

            Shape3D shape = new Shape3D();
            effect.shape3D = shape;
            Bevel bevel = new Bevel();
            shape.bevelTop = bevel;
            bevel.width = 4D;
            bevel.height = 1.3D;
            shape.contourColor = new Color(192, 192, 192);
            shape.contourWidth = 0.5D;
            return effect;
        }
    },
    // 19
    RelaxedPerspective_White {
        @Override public Effect getEffect() {
            if (effect != null) return effect;
            effect = new Effect();

            Fill.SolidFill fill = new Fill.SolidFill();
            fill.color = Color.WHITE;
            fill.shade = 85;
            effect.fill = fill;

            Shadow shadow = new Shadow();
            shadow.color = Color.BLACK;
            shadow.alpha = 80;
            shadow.blur = 4.5D;
            shadow.direction = 126;
            shadow.dist = 3D;
            shadow.angle = Angle.TOP_LEFT;
            shadow.sy = 98D;
            shadow.kx = 1.83D;
            shadow.ky = 3.33D;
            effect.shadow = shadow;

            effect.geometry = ShapeType.rect;

            Outline ln = new Outline();
            effect.outline = ln;
            ln.width = 8D;
            ln.color = new Color(253, 253, 253);
            ln.cap = Cap.square;
            ln.joinType = JoinType.miter;
            ln.miterLimit = 800D;

            Scene3D scene = new Scene3D();
            effect.scene3D = scene;
            Camera camera = new Camera();
            camera.presetCamera = PresetCamera.perspectiveRelaxed;
            camera.latitude = 316D;
            scene.camera = camera;
            LightRig lightRig = new LightRig();
            lightRig.rig = Rig.twoPt;
            lightRig.angle = Angle.TOP;
            lightRig.revolution = 120D;
            scene.lightRig = lightRig;

            Shape3D shape = new Shape3D();
            effect.shape3D = shape;
            shape.material = Material.matte;
            Bevel bevel = new Bevel();
            shape.bevelTop = bevel;
            bevel.width = 1.8D;
            bevel.height = 1D;
            shape.contourColor = Color.WHITE;
            return effect;
        }
    },
    // 20
    SoftEdgeOval_White {
        @Override public Effect getEffect() {
            if (effect != null) return effect;
            effect = new Effect();

            effect.geometry = ShapeType.ellipse;
            effect.softEdges = 8.86D;
            return effect;
        }
    },
    // 21
    BevelRectangle {
        @Override public Effect getEffect() {
            if (effect != null) return effect;
            effect = new Effect();

            Shadow shadow = new Shadow();
            shadow.color = Color.BLACK;
            shadow.alpha = 60;
            shadow.blur = 6D;
            shadow.direction = 130;
            shadow.dist = 3D;
            shadow.angle = Angle.TOP_LEFT;
            effect.shadow = shadow;

            effect.geometry = ShapeType.roundRect;
            effect.geometryAdjustValueList = Collections.singletonList(Tuple2.of("adj", "val 16667"));

            Scene3D scene = new Scene3D();
            effect.scene3D = scene;
            Camera camera = new Camera();
            camera.presetCamera = PresetCamera.orthographicFront;
            scene.camera = camera;
            LightRig lightRig = new LightRig();
            lightRig.rig = Rig.contrasting;
            lightRig.angle = Angle.TOP;
            lightRig.revolution = 70D;
            scene.lightRig = lightRig;

            Shape3D shape = new Shape3D();
            effect.shape3D = shape;
            shape.material = Material.plastic;
            Bevel bevel = new Bevel();
            shape.bevelTop = bevel;
            bevel.width = 30D;
            bevel.height = 9D;
            bevel.prst = PresetBevel.relaxedInset;
            shape.contourColor = new Color(150, 150, 150);
            return effect;
        }
    },
    // 22
    BevelPerspective {
        @Override public Effect getEffect() {
            if (effect != null) return effect;
            effect = new Effect();

            Shadow shadow = new Shadow();
            shadow.color = Color.BLACK;
            shadow.alpha = 70;
            shadow.blur = 12D;
            shadow.direction = 15;
            shadow.dist =0.94D;
            shadow.angle = Angle.TOP_LEFT;
            shadow.sy = 98D;
            shadow.kx = 1.83D;
            shadow.ky = 3.33D;
            effect.shadow = shadow;

            effect.geometry = ShapeType.roundRect;
            effect.geometryAdjustValueList = Collections.singletonList(Tuple2.of("adj", "val 16667"));

            Scene3D scene = new Scene3D();
            effect.scene3D = scene;
            Camera camera = new Camera();
            camera.presetCamera = PresetCamera.perspectiveRelaxed;
            camera.latitude = 330D;
            camera.longitude = 20D;
            camera.revolution = 347D;
            scene.camera = camera;
            LightRig lightRig = new LightRig();
            lightRig.rig = Rig.threePt;
            lightRig.angle = Angle.TOP;
            scene.lightRig = lightRig;

            Shape3D shape = new Shape3D();
            effect.shape3D = shape;
            shape.material = Material.matte;
            Bevel bevel = new Bevel();
            shape.bevelTop = bevel;
            bevel.width = 8D;
            bevel.height = 8D;
            shape.contourColor = new Color(150, 150, 150);
            shape.contourWidth = 0.5D;
            return effect;
        }
    },
    // 23
    ReflectedPerspectiveRight {
        @Override public Effect getEffect() {
            if (effect != null) return effect;
            effect = new Effect();

            Reflection reflection = new Reflection();
            reflection.blur = 1D;
            reflection.alpha = 70;
            reflection.size = 30D;
            reflection.dist = 0.4D;
            effect.reflection = reflection;

            effect.geometry = ShapeType.rect;

            Scene3D scene = new Scene3D();
            effect.scene3D = scene;
            Camera camera = new Camera();
            camera.presetCamera = PresetCamera.perspectiveContrastingLeftFacing;
            camera.latitude = 5D;
            camera.longitude = 330D;
            scene.camera = camera;
            LightRig lightRig = new LightRig();
            lightRig.rig = Rig.threePt;
            lightRig.angle = Angle.TOP;
            lightRig.revolution = 45D;
            scene.lightRig = lightRig;

            Shape3D shape = new Shape3D();
            effect.shape3D = shape;
            Bevel bevel = new Bevel();
            shape.bevelTop = bevel;
            bevel.width = 5D;
            bevel.height = 4D;
            return effect;
        }
    },
    // 24
    BevelPerspectiveLeft_White {
        @Override public Effect getEffect() {
            if (effect != null) return effect;
            effect = new Effect();

            Fill.SolidFill fill = new Fill.SolidFill();
            fill.color = Color.WHITE;
            fill.shade = 85;
            effect.fill = fill;

            Shadow shadow = new Shadow();
            shadow.color = Color.BLACK;
            shadow.alpha = 67;
            shadow.blur = 2.85D;
            shadow.direction = 190;
            shadow.dist = 1D;
            shadow.angle = Angle.TOP_LEFT;
            effect.shadow = shadow;

            effect.geometry = ShapeType.rect;

            Outline ln = new Outline();
            effect.outline = ln;
            ln.width = 15D;
            ln.color = Color.WHITE;
            ln.cap = Cap.round;

            Scene3D scene = new Scene3D();
            effect.scene3D = scene;
            Camera camera = new Camera();
            camera.presetCamera = PresetCamera.perspectiveContrastingLeftFacing;
            camera.latitude = 9D;
            camera.longitude = 35D;
            scene.camera = camera;
            LightRig lightRig = new LightRig();
            lightRig.rig = Rig.soft;
            lightRig.angle = Angle.TOP;
            scene.lightRig = lightRig;

            Shape3D shape = new Shape3D();
            effect.shape3D = shape;
            shape.material = Material.matte;
            Bevel bevel = new Bevel();
            shape.bevelTop = bevel;
            bevel.width = 5D;
            bevel.height = 4D;
            shape.contourColor = new Color(192, 192, 192);
            shape.contourWidth = 1D;
            return effect;
        }
    },
    // 25
    ReflectedBevel_Black {
        @Override public Effect getEffect() {
            if (effect != null) return effect;
            effect = new Effect();

            Fill.SolidFill fill = new Fill.SolidFill();
            fill.color = Color.WHITE;
            effect.fill = fill;

            Reflection reflection = new Reflection();
            reflection.blur = 1D;
            reflection.alpha = 72;
            reflection.size = 28D;
            reflection.dist = 0.4D;
            effect.reflection = reflection;

            effect.geometry = ShapeType.roundRect;
            effect.geometryAdjustValueList = Collections.singletonList(Tuple2.of("adj", "val 4167"));

            Outline ln = new Outline();
            effect.outline = ln;
            ln.width = 6D;
            ln.color = new Color(41, 41, 41);
            ln.cap = Cap.square;
            ln.joinType = JoinType.miter;
            ln.miterLimit = 800D;

            Scene3D scene = new Scene3D();
            effect.scene3D = scene;
            Camera camera = new Camera();
            camera.presetCamera = PresetCamera.orthographicFront;
            scene.camera = camera;
            LightRig lightRig = new LightRig();
            lightRig.rig = Rig.threePt;
            lightRig.angle = Angle.TOP;
            lightRig.revolution = 45D;
            scene.lightRig = lightRig;

            Shape3D shape = new Shape3D();
            effect.shape3D = shape;
            Bevel bevel = new Bevel();
            shape.bevelTop = bevel;
            bevel.height = 3D;
            shape.contourColor = new Color(192, 192, 192);
            return effect;
        }
    },
    // 26
    ReflectedBevel_White {
        @Override public Effect getEffect() {
            if (effect != null) return effect;
            effect = new Effect();

            Fill.SolidFill fill = new Fill.SolidFill();
            fill.color = Color.WHITE;
            effect.fill = fill;

            Reflection reflection = new Reflection();
            reflection.blur = 1D;
            reflection.alpha = 67;
            reflection.size = 28D;
            reflection.dist = 0.4D;
            effect.reflection = reflection;

            effect.geometry = ShapeType.roundRect;
            effect.geometryAdjustValueList = Collections.singletonList(Tuple2.of("adj", "val 4167"));

            Outline ln = new Outline();
            effect.outline = ln;
            ln.width = 6D;
            ln.color = new Color(234, 234, 234);
            ln.cap = Cap.square;
            ln.joinType = JoinType.miter;
            ln.miterLimit = 800D;

            Scene3D scene = new Scene3D();
            effect.scene3D = scene;
            Camera camera = new Camera();
            camera.presetCamera = PresetCamera.orthographicFront;
            scene.camera = camera;
            LightRig lightRig = new LightRig();
            lightRig.rig = Rig.threePt;
            lightRig.angle = Angle.TOP;
            lightRig.revolution = 45D;
            scene.lightRig = lightRig;

            Shape3D shape = new Shape3D();
            effect.shape3D = shape;
            Bevel bevel = new Bevel();
            shape.bevelTop = bevel;
            bevel.height = 3D;
            bevel.width = 6D;
            shape.contourColor = new Color(192, 192, 192);
            shape.contourWidth = 0.5D;
            return effect;
        }
    },
    // 27
    MetalRoundedRectangle {
        @Override public Effect getEffect() {
            if (effect != null) return effect;
            effect = new Effect();

            Shadow shadow = new Shadow();
            shadow.color = Color.BLACK;
            shadow.alpha = 55;
            shadow.blur = 8D;
            shadow.direction = 120;
            shadow.dist = 4D;
            shadow.angle = Angle.TOP_LEFT;
            effect.shadow = shadow;

            effect.geometry = ShapeType.roundRect;
            effect.geometryAdjustValueList = Collections.singletonList(Tuple2.of("adj", "val 11111"));

            Outline ln = new Outline();
            effect.outline = ln;
            ln.width = 15D;
            ln.color = new Color(200, 198, 189);
            ln.cap = Cap.round;
            ln.dash = DashPattern.solid;

            Scene3D scene = new Scene3D();
            effect.scene3D = scene;
            Camera camera = new Camera();
            camera.presetCamera = PresetCamera.perspectiveFront;
            camera.fov = 90D;
            scene.camera = camera;
            LightRig lightRig = new LightRig();
            lightRig.rig = Rig.threePt;
            lightRig.angle = Angle.TOP;
            lightRig.revolution = 320D;
            scene.lightRig = lightRig;

            Shape3D shape = new Shape3D();
            effect.shape3D = shape;
            Bevel bevel = new Bevel();
            shape.bevelTop = bevel;
            bevel.width = 24D;
            bevel.height = 12D;
            bevel.prst = PresetBevel.hardEdge;
            shape.extrusionColor = Color.WHITE;
            shape.extrusionHeight = 2D;
            return effect;
        }
    },
    // 28
    MetalOval {
        @Override public Effect getEffect() {
            if (effect != null) return effect;
            effect = new Effect();

            Shadow shadow = new Shadow();
            shadow.color = Color.BLACK;
            shadow.blur = 10D;
            shadow.angle = Angle.BOTTOM_LEFT;
            effect.shadow = shadow;

            effect.geometry = ShapeType.ellipse;

            Outline ln = new Outline();
            effect.outline = ln;
            ln.width = 15D;
            ln.color = new Color(200, 198, 189);
            ln.cap = Cap.round;
            ln.dash = DashPattern.solid;

            Scene3D scene = new Scene3D();
            effect.scene3D = scene;
            Camera camera = new Camera();
            camera.presetCamera = PresetCamera.perspectiveFront;
            camera.fov = 90D;
            scene.camera = camera;
            LightRig lightRig = new LightRig();
            lightRig.rig = Rig.threePt;
            lightRig.angle = Angle.TOP;
            lightRig.revolution = 320D;
            scene.lightRig = lightRig;

            Shape3D shape = new Shape3D();
            effect.shape3D = shape;
            Bevel bevel = new Bevel();
            shape.bevelTop = bevel;
            bevel.width = 24D;
            bevel.height = 12D;
            bevel.prst = PresetBevel.hardEdge;
            shape.extrusionColor = Color.BLACK;
            shape.extrusionHeight = 2D;
            return effect;
        }
    }
    ;

    protected Effect effect;
}

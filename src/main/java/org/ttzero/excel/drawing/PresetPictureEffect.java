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

import org.ttzero.excel.entity.Picture;

import java.awt.Color;
import java.util.Arrays;
import java.util.Collections;


/**
 * Preset Picture Effects
 *
 * @author guanquan.wang at 2023-07-25 09:59
 */
public enum PresetPictureEffect {
    // 1
    SimpleFrame_White {
        @Override public void preset(Picture pict) {
            PictureEffect effect = new PictureEffect();
            pict.effect = effect;

            Fill.SolidFill fill = new Fill.SolidFill();
            fill.color = Color.WHITE;
            fill.shade = 85;
            effect.fill = fill;

            Outline ln = new Outline();
            effect.outline = ln;
            ln.width = 7;
            ln.color = Color.WHITE;
            ln.cap = Outline.Cap.SQUARE;
            ln.joinType = Outline.JoinType.miter;
            ln.miterLimit = 800;

            Shadow shadow = new Shadow();
            effect.shadow = shadow;
            shadow.color = Color.BLACK;
            shadow.alpha = 60;
            shadow.blur = 4.33;
            shadow.direction = 90;
            shadow.dist = 1.42;
            shadow.angle = Angle.TOP_LEFT;

            effect.geometry = ShapeType.rect;

            Scene3D scene = new Scene3D();
            effect.scene3D = scene;
            Camera camera = new Camera();
            camera.presetCamera = Camera.PresetCamera.orthographicFront;
            scene.camera = camera;
            LightRig lightRig = new LightRig();
            lightRig.rig = LightRig.Rig.twoPt;
            lightRig.angle = Angle.TOP;
            lightRig.revolution = 120;
            scene.lightRig = lightRig;

            Shape3D shape = new Shape3D();
            effect.shape3D = shape;
            Bevel bevel = new Bevel();
            shape.bevelTop = bevel;
            bevel.width = 2;
            bevel.height = 1.5;
            shape.contourColor = Color.WHITE;
        }
    },
    // 2
    BeveledMatte_White {
        @Override public void preset(Picture pict) {
            PictureEffect effect = new PictureEffect();
            pict.effect = effect;

            Fill.SolidFill fill = new Fill.SolidFill();
            fill.color = Color.WHITE;
            fill.shade = 85;
            effect.fill = fill;

            Outline ln = new Outline();
            effect.outline = ln;
            ln.width = 15;
            ln.color = Color.WHITE;
            ln.cap = Outline.Cap.ROUND;

            Shadow shadow = new Shadow();
            effect.shadow = shadow;
            shadow.color = Color.BLACK;
            shadow.alpha = 59;
            shadow.blur = 3.94;
            shadow.angle = Angle.TOP_LEFT;

            effect.geometry = ShapeType.rect;

            Scene3D scene = new Scene3D();
            effect.scene3D = scene;
            Camera camera = new Camera();
            camera.presetCamera = Camera.PresetCamera.orthographicFront;
            scene.camera = camera;
            LightRig lightRig = new LightRig();
            lightRig.rig = LightRig.Rig.twoPt;
            lightRig.angle = Angle.TOP;
            lightRig.revolution = 130;
            scene.lightRig = lightRig;

            Shape3D shape = new Shape3D();
            effect.shape3D = shape;
            Bevel bevel = new Bevel();
            shape.bevelTop = bevel;
            bevel.width = 4;
            bevel.height = 1.3;
            shape.contourColor = new Color(192, 192, 192);
            shape.contourWidth = 0.5;
        }
    },
    // 3
    MetalFrame {
        @Override public void preset(Picture pict) {
            PictureEffect effect = new PictureEffect();
            pict.effect = effect;
            Outline ln = new Outline();
            effect.outline = ln;
            ln.width = 15;
            ln.color = new Color(200, 198, 189);
            ln.cap = Outline.Cap.SQUARE;
            ln.dash = Outline.DashPattern.solid;
            ln.joinType = Outline.JoinType.miter;
            ln.miterLimit = 800;

            Shadow shadow = new Shadow();
            effect.shadow = shadow;
            shadow.color = Color.BLACK;
            shadow.alpha = 57;
            shadow.blur = 20;
            shadow.angle = Angle.BOTTOM_LEFT;

            effect.geometry = ShapeType.rect;

            Scene3D scene = new Scene3D();
            effect.scene3D = scene;
            Camera camera = new Camera();
            camera.presetCamera = Camera.PresetCamera.perspectiveFront;
            camera.fov = 90;
            scene.camera = camera;
            LightRig lightRig = new LightRig();
            lightRig.rig = LightRig.Rig.threePt;
            lightRig.angle = Angle.TOP;
            lightRig.revolution = 35;
            scene.lightRig = lightRig;

            Shape3D shape = new Shape3D();
            effect.shape3D = shape;
            Bevel bevel = new Bevel();
            shape.bevelTop = bevel;
            bevel.width = 24;
            bevel.height = 12;
            bevel.prst = Bevel.BevelPresetType.hardEdge;
            shape.extrusionColor = Color.BLACK;
            shape.extrusionHeight = 2;
        }
    },
    // 4
    DropShadowRectangle {
        @Override public void preset(Picture pict) {
            PictureEffect effect = new PictureEffect();
            pict.effect = effect;
            Shadow shadow = new Shadow();
            shadow.color = new Color(51, 51, 51);
            shadow.alpha = 35;
            shadow.direction = 45;
            shadow.blur = 23;
            shadow.dist = 11;
            shadow.angle = Angle.TOP_LEFT;
            effect.shadow = shadow;

            effect.geometry = ShapeType.rect;
        }
    },
    // 5
    ReflectedRoundedRectangle {
        @Override public void preset(Picture pict) {
            PictureEffect effect = new PictureEffect();
            pict.effect = effect;

            Fill.SolidFill fill = new Fill.SolidFill();
            fill.color = Color.WHITE;
            fill.shade = 85;
            effect.fill = fill;

            Reflection reflection = new Reflection();
            reflection.blur = 1;
            reflection.alpha = 62;
            reflection.size = 28;
            reflection.dist = 0.4D;
            effect.reflection = reflection;

            effect.geometry = ShapeType.roundRect;
            effect.geometryAdjustValueList = Collections.singletonList(new Guide("adj", "val 8594"));
        }
    },
    // 6
    SoftEdgeRectangle {
        @Override public void preset(Picture pict) {
            PictureEffect effect = new PictureEffect();
            pict.effect = effect;
            effect.softEdges = 8.86D;
            effect.geometry = ShapeType.rect;
        }
    },
    // 7
    DoubleFrame_Black {
        @Override public void preset(Picture pict) {
            PictureEffect effect = new PictureEffect();
            pict.effect = effect;
            Outline ln = new Outline();
            effect.outline = ln;
            ln.width = 18;
            ln.color = Color.BLACK;
            ln.cap = Outline.Cap.SQUARE;
            ln.cmpd = Outline.CompoundType.thickThin;
            ln.dash = Outline.DashPattern.solid;
            ln.joinType = Outline.JoinType.miter;
            ln.miterLimit = 800;

            Shadow shadow = new Shadow();
            effect.innerShadow = shadow;
            shadow.color = Color.BLACK;
            shadow.blur = 6;

            effect.geometry = ShapeType.rect;
        }
    },
    // 8
    ThickMatte_Black {
        @Override public void preset(Picture pict) {
            PictureEffect effect = new PictureEffect();
            pict.effect = effect;

            Fill.SolidFill fill = new Fill.SolidFill();
            fill.color = Color.BLACK;
            fill.shade = 95;
            effect.fill = fill;

            Outline ln = new Outline();
            effect.outline = ln;
            ln.width = 35;
            ln.color = Color.BLACK;
            ln.cap = Outline.Cap.SQUARE;
            ln.joinType = Outline.JoinType.miter;
            ln.miterLimit = 800;

            Shadow shadow = new Shadow();
            effect.shadow = shadow;
            shadow.color = Color.BLACK;
            shadow.alpha = 60;
            shadow.blur = 20;
            shadow.direction = 45;
            shadow.dist = 15;
            shadow.angle = Angle.BOTTOM_LEFT;
            shadow.sy = 90;

            effect.geometry = ShapeType.rect;
        }
    },
    // 9
    SimpleFrame_Black {
        @Override public void preset(Picture pict) {
            PictureEffect effect = new PictureEffect();
            pict.effect = effect;
            Outline ln = new Outline();
            effect.outline = ln;
            ln.width = 3;
            ln.color = Color.BLACK;
            ln.cap = Outline.Cap.SQUARE;
            ln.dash = Outline.DashPattern.solid;
            ln.joinType = Outline.JoinType.miter;
            ln.miterLimit = 800;

            Shadow shadow = new Shadow();
            effect.shadow = shadow;
            shadow.color = Color.BLACK;
            shadow.alpha = 57;
            shadow.blur = 4;
            shadow.direction = 45;
            shadow.dist = 3;
            shadow.angle = Angle.TOP_LEFT;

            effect.geometry = ShapeType.rect;
        }
    },
    // 10
    BeveledOval_Black {
        @Override public void preset(Picture pict) {
            PictureEffect effect = new PictureEffect();
            pict.effect = effect;
            Outline ln = new Outline();
            effect.outline = ln;
            ln.width = 5;
            ln.color = new Color(51, 51, 51);
            ln.cap = Outline.Cap.ROUND;

            Shadow shadow = new Shadow();
            effect.shadow = shadow;
            shadow.color = Color.BLACK;
            shadow.alpha = 78;
            shadow.blur = 30;
            shadow.direction = 90;
            shadow.dist = 23;
            shadow.sx = -80;
            shadow.sy = -18;

            effect.geometry = ShapeType.ellipse;

            Scene3D scene = new Scene3D();
            effect.scene3D = scene;
            Camera camera = new Camera();
            camera.presetCamera = Camera.PresetCamera.orthographicFront;
            scene.camera = camera;
            LightRig lightRig = new LightRig();
            lightRig.rig = LightRig.Rig.contrasting;
            lightRig.angle = Angle.TOP;
            lightRig.revolution = 50;
            scene.lightRig = lightRig;

            Shape3D shape = new Shape3D();
            effect.shape3D = shape;
            Bevel bevel = new Bevel();
            shape.bevelTop = bevel;
            bevel.width = 7.5;
            bevel.height = 2.5;
            shape.contourColor = new Color(51, 51, 51);
            shape.contourWidth = 0.6;
        }
    },
    // 11
    CompoundFrame_Black {
        @Override public void preset(Picture pict) {
            PictureEffect effect = new PictureEffect();
            pict.effect = effect;
            Outline ln = new Outline();
            effect.outline = ln;
            ln.width = 7;
            ln.color = Color.BLACK;
            ln.cap = Outline.Cap.SQUARE;
            ln.dash = Outline.DashPattern.solid;
            ln.cmpd = Outline.CompoundType.thickThin;
            ln.joinType = Outline.JoinType.miter;
            ln.miterLimit = 800;

            Shadow shadow = new Shadow();
            effect.innerShadow = shadow;
            shadow.color = Color.BLACK;
            shadow.blur = 6;

            effect.geometry = ShapeType.rect;
        }
    },
    // 12
    ModerateFrame_Black {
        @Override public void preset(Picture pict) {
            PictureEffect effect = new PictureEffect();
            pict.effect = effect;
            Outline ln = new Outline();
            effect.outline = ln;
            ln.width = 10;
            ln.color = Color.BLACK;
            ln.cap = Outline.Cap.SQUARE;
            ln.joinType = Outline.JoinType.miter;
            ln.miterLimit = 800;

            Shadow shadow = new Shadow();
            effect.shadow = shadow;
            shadow.color = Color.BLACK;
            shadow.alpha = 60;
            shadow.blur = 4.5;
            shadow.direction = 45;
            shadow.dist = 4;
            shadow.angle = Angle.TOP_LEFT;

            effect.geometry = ShapeType.rect;
        }
    },
    // 13
    CenterShadowRectangle {
        @Override public void preset(Picture pict) {
            PictureEffect effect = new PictureEffect();
            pict.effect = effect;
            Shadow shadow = new Shadow();
            shadow.color = Color.BLACK;
            shadow.alpha = 30;
            shadow.blur = 15;
            shadow.angle = Angle.TOP_LEFT;
            effect.shadow = shadow;

            effect.geometry = ShapeType.rect;
        }
    },
    // 14
    RoundedDiagonalCorner_White {
        @Override public void preset(Picture pict) {
            PictureEffect effect = new PictureEffect();
            pict.effect = effect;
            Shadow shadow = new Shadow();
            shadow.color = Color.BLACK;
            shadow.alpha = 57;
            shadow.blur = 20;
            shadow.angle = Angle.TOP_LEFT;
            effect.shadow = shadow;

            effect.geometry = ShapeType.round2DiagRect;
            effect.geometryAdjustValueList = Arrays.asList(new Guide("adj1", "val 16667"), new Guide("adj2", "val 0"));

            Outline ln = new Outline();
            effect.outline = ln;
            ln.width = 7;
            ln.color = Color.WHITE;
            ln.cap = Outline.Cap.SQUARE;
            ln.joinType = Outline.JoinType.miter;
            ln.miterLimit = 800;
        }
    },
    // 15
    SnipDiagonalCorner_White {
        @Override public void preset(Picture pict) {
            PictureEffect effect = new PictureEffect();
            pict.effect = effect;

            Fill.SolidFill fill = new Fill.SolidFill();
            fill.color = Color.WHITE;
            fill.shade = 85;
            effect.fill = fill;

            Shadow shadow = new Shadow();
            shadow.color = Color.BLACK;
            shadow.alpha = 55;
            shadow.blur = 7;
            shadow.angle = Angle.TOP_LEFT;
            effect.shadow = shadow;

            effect.geometry = ShapeType.snip2DiagRect;

            Outline ln = new Outline();
            effect.outline = ln;
            ln.width = 7;
            ln.color = Color.WHITE;
            ln.cap = Outline.Cap.SQUARE;
            ln.joinType = Outline.JoinType.miter;
            ln.miterLimit = 800;

            Scene3D scene = new Scene3D();
            effect.scene3D = scene;
            Camera camera = new Camera();
            camera.presetCamera = Camera.PresetCamera.orthographicFront;
            scene.camera = camera;
            LightRig lightRig = new LightRig();
            lightRig.rig = LightRig.Rig.twoPt;
            lightRig.angle = Angle.TOP;
            lightRig.revolution = 120;
            scene.lightRig = lightRig;

            Shape3D shape = new Shape3D();
            effect.shape3D = shape;
            Bevel bevel = new Bevel();
            shape.bevelTop = bevel;
            bevel.width = 2;
            bevel.height = 1.5;
            shape.contourColor = Color.WHITE;
        }
    },
    // 16
    ModerateFrame_White {
        @Override public void preset(Picture pict) {
            PictureEffect effect = new PictureEffect();
            pict.effect = effect;

            Fill.SolidFill fill = new Fill.SolidFill();
            fill.color = Color.WHITE;
            fill.shade = 85;
            effect.fill = fill;

            Shadow shadow = new Shadow();
            shadow.color = Color.BLACK;
            shadow.alpha = 60;
            shadow.blur = 4.33;
            shadow.direction = 90;
            shadow.dist = 1.4;
            shadow.angle = Angle.TOP_LEFT;
            effect.shadow = shadow;

            effect.geometry = ShapeType.rect;

            Outline ln = new Outline();
            effect.outline = ln;
            ln.width = 15;
            ln.color = Color.WHITE;
            ln.cap = Outline.Cap.SQUARE;
            ln.joinType = Outline.JoinType.miter;
            ln.miterLimit = 800;

            Scene3D scene = new Scene3D();
            effect.scene3D = scene;
            Camera camera = new Camera();
            camera.presetCamera = Camera.PresetCamera.orthographicFront;
            scene.camera = camera;
            LightRig lightRig = new LightRig();
            lightRig.rig = LightRig.Rig.twoPt;
            lightRig.angle = Angle.TOP;
            lightRig.revolution = 120;
            scene.lightRig = lightRig;

            Shape3D shape = new Shape3D();
            effect.shape3D = shape;
            Bevel bevel = new Bevel();
            shape.bevelTop = bevel;
            bevel.width = 2;
            bevel.height = 1.5;
            shape.contourColor = Color.WHITE;
        }
    },
    // 17
    Rotated_White {
        @Override public void preset(Picture pict) {
            PictureEffect effect = new PictureEffect();
            pict.effect = effect;

            Fill.SolidFill fill = new Fill.SolidFill();
            fill.color = Color.WHITE;
            fill.shade = 85;
            effect.fill = fill;

            Shadow shadow = new Shadow();
            shadow.color = Color.BLACK;
            shadow.alpha = 70;
            shadow.blur = 5.12;
            shadow.direction = 215;
            shadow.dist = 4;
            shadow.angle = Angle.TOP_LEFT;
            shadow.kx = 3.25;
            shadow.ky = 2.42;
            effect.shadow = shadow;

            effect.geometry = ShapeType.rect;

            Outline ln = new Outline();
            effect.outline = ln;
            ln.width = 15;
            ln.color = Color.WHITE;
            ln.cap = Outline.Cap.SQUARE;
            ln.joinType = Outline.JoinType.miter;
            ln.miterLimit = 800;

            Scene3D scene = new Scene3D();
            effect.scene3D = scene;
            Camera camera = new Camera();
            camera.presetCamera = Camera.PresetCamera.orthographicFront;
            camera.revolution = 6;
            scene.camera = camera;
            LightRig lightRig = new LightRig();
            lightRig.rig = LightRig.Rig.twoPt;
            lightRig.angle = Angle.TOP;
            lightRig.revolution = 120;
            scene.lightRig = lightRig;

            Shape3D shape = new Shape3D();
            effect.shape3D = shape;
            Bevel bevel = new Bevel();
            shape.bevelTop = bevel;
            bevel.width = 2;
            bevel.height = 1.5;
            shape.contourColor = new Color(150, 150, 150);
            shape.contourWidth = 1;
        }
    },
    // 18
    PerspectiveShadow_White {
        @Override public void preset(Picture pict) {
            PictureEffect effect = new PictureEffect();
            pict.effect = effect;
            Shadow shadow = new Shadow();
            shadow.color = Color.BLACK;
            shadow.alpha = 80;
            shadow.blur = 6;
            shadow.direction = 175;
            shadow.dist = 7.5;
            shadow.angle = Angle.BOTTOM_RIGHT;
            shadow.kx = 15;
            shadow.sx = 97;
            shadow.sy = 23;
            effect.shadow = shadow;

            effect.geometry = ShapeType.rect;

            Outline ln = new Outline();
            effect.outline = ln;
            ln.width = 10;
            ln.color = Color.WHITE;
            ln.cap = Outline.Cap.ROUND;

            Scene3D scene = new Scene3D();
            effect.scene3D = scene;
            Camera camera = new Camera();
            camera.presetCamera = Camera.PresetCamera.orthographicFront;
            scene.camera = camera;
            LightRig lightRig = new LightRig();
            lightRig.rig = LightRig.Rig.twoPt;
            lightRig.angle = Angle.TOP;
            lightRig.revolution = 130;
            scene.lightRig = lightRig;

            Shape3D shape = new Shape3D();
            effect.shape3D = shape;
            Bevel bevel = new Bevel();
            shape.bevelTop = bevel;
            bevel.width = 4;
            bevel.height = 1.3;
            shape.contourColor = new Color(192, 192, 192);
            shape.contourWidth = 0.5;
        }
    },
    // 19
    RelaxedPerspective_White {
        @Override public void preset(Picture pict) {
            PictureEffect effect = new PictureEffect();
            pict.effect = effect;

            Fill.SolidFill fill = new Fill.SolidFill();
            fill.color = Color.WHITE;
            fill.shade = 85;
            effect.fill = fill;

            Shadow shadow = new Shadow();
            shadow.color = Color.BLACK;
            shadow.alpha = 80;
            shadow.blur = 4.5;
            shadow.direction = 126;
            shadow.dist = 3;
            shadow.angle = Angle.TOP_LEFT;
            shadow.sy = 98;
            shadow.kx = 1.83;
            shadow.ky = 3.33;
            effect.shadow = shadow;

            effect.geometry = ShapeType.rect;

            Outline ln = new Outline();
            effect.outline = ln;
            ln.width = 8;
            ln.color = new Color(253, 253, 253);
            ln.cap = Outline.Cap.SQUARE;
            ln.joinType = Outline.JoinType.miter;
            ln.miterLimit = 800;

            Scene3D scene = new Scene3D();
            effect.scene3D = scene;
            Camera camera = new Camera();
            camera.presetCamera = Camera.PresetCamera.perspectiveRelaxed;
            camera.latitude = 316;
            scene.camera = camera;
            LightRig lightRig = new LightRig();
            lightRig.rig = LightRig.Rig.twoPt;
            lightRig.angle = Angle.TOP;
            lightRig.revolution = 120;
            scene.lightRig = lightRig;

            Shape3D shape = new Shape3D();
            effect.shape3D = shape;
            shape.material = Shape3D.Material.matte;
            Bevel bevel = new Bevel();
            shape.bevelTop = bevel;
            bevel.width = 1.8;
            bevel.height = 1;
            shape.contourColor = Color.WHITE;
        }
    },
    // 20
    SoftEdgeOval_White {
        @Override public void preset(Picture pict) {
            PictureEffect effect = new PictureEffect();
            pict.effect = effect;
            effect.geometry = ShapeType.ellipse;
            effect.softEdges = 8.86;
        }
    },
    // 21
    BevelRectangle {
        @Override public void preset(Picture pict) {
            PictureEffect effect = new PictureEffect();
            pict.effect = effect;
            Shadow shadow = new Shadow();
            shadow.color = Color.BLACK;
            shadow.alpha = 60;
            shadow.blur = 6;
            shadow.direction = 130;
            shadow.dist = 3;
            shadow.angle = Angle.TOP_LEFT;
            effect.shadow = shadow;

            effect.geometry = ShapeType.roundRect;
            effect.geometryAdjustValueList = Collections.singletonList(new Guide("adj", "val 16667"));

            Scene3D scene = new Scene3D();
            effect.scene3D = scene;
            Camera camera = new Camera();
            camera.presetCamera = Camera.PresetCamera.orthographicFront;
            scene.camera = camera;
            LightRig lightRig = new LightRig();
            lightRig.rig = LightRig.Rig.contrasting;
            lightRig.angle = Angle.TOP;
            lightRig.revolution = 70;
            scene.lightRig = lightRig;

            Shape3D shape = new Shape3D();
            effect.shape3D = shape;
            shape.material = Shape3D.Material.plastic;
            Bevel bevel = new Bevel();
            shape.bevelTop = bevel;
            bevel.width = 30;
            bevel.height = 9;
            bevel.prst = Bevel.BevelPresetType.relaxedInset;
            shape.contourColor = new Color(150, 150, 150);
        }
    },
    // 22
    BevelPerspective {
        @Override public void preset(Picture pict) {
            PictureEffect effect = new PictureEffect();
            pict.effect = effect;
            Shadow shadow = new Shadow();
            shadow.color = Color.BLACK;
            shadow.alpha = 70;
            shadow.blur = 12;
            shadow.direction = 15;
            shadow.dist =0.94;
            shadow.angle = Angle.TOP_LEFT;
            shadow.sy = 98;
            shadow.kx = 1.83;
            shadow.ky = 3.33;
            effect.shadow = shadow;

            effect.geometry = ShapeType.roundRect;
            effect.geometryAdjustValueList = Collections.singletonList(new Guide("adj", "val 16667"));

            Scene3D scene = new Scene3D();
            effect.scene3D = scene;
            Camera camera = new Camera();
            camera.presetCamera = Camera.PresetCamera.perspectiveRelaxed;
            camera.latitude = 330;
            camera.longitude = 20;
            camera.revolution = 347;
            scene.camera = camera;
            LightRig lightRig = new LightRig();
            lightRig.rig = LightRig.Rig.threePt;
            lightRig.angle = Angle.TOP;
            scene.lightRig = lightRig;

            Shape3D shape = new Shape3D();
            effect.shape3D = shape;
            shape.material = Shape3D.Material.matte;
            Bevel bevel = new Bevel();
            shape.bevelTop = bevel;
            bevel.width = 8;
            bevel.height = 8;
            shape.contourColor = new Color(150, 150, 150);
            shape.contourWidth = 0.5;
        }
    },
    // 23
    ReflectedperpsectiveRight {
        @Override public void preset(Picture pict) {
            PictureEffect effect = new PictureEffect();
            pict.effect = effect;
            Reflection reflection = new Reflection();
            reflection.blur = 1;
            reflection.alpha = 70;
            reflection.size = 30;
            reflection.dist = 0.4D;
            effect.reflection = reflection;

            effect.geometry = ShapeType.rect;

            Scene3D scene = new Scene3D();
            effect.scene3D = scene;
            Camera camera = new Camera();
            camera.presetCamera = Camera.PresetCamera.perspectiveContrastingLeftFacing;
            camera.latitude = 5;
            camera.longitude = 330;
            camera.revolution = 0;
            scene.camera = camera;
            LightRig lightRig = new LightRig();
            lightRig.rig = LightRig.Rig.threePt;
            lightRig.angle = Angle.TOP;
            lightRig.revolution = 45;
            scene.lightRig = lightRig;

            Shape3D shape = new Shape3D();
            effect.shape3D = shape;
            Bevel bevel = new Bevel();
            shape.bevelTop = bevel;
            bevel.width = 5;
            bevel.height = 4;
        }
    },
    // 24
    BevelPerspectiveLeft_White {
        @Override public void preset(Picture pict) {
            PictureEffect effect = new PictureEffect();
            pict.effect = effect;

            Fill.SolidFill fill = new Fill.SolidFill();
            fill.color = Color.WHITE;
            fill.shade = 85;
            effect.fill = fill;

            Shadow shadow = new Shadow();
            shadow.color = Color.BLACK;
            shadow.alpha = 67;
            shadow.blur = 2.85;
            shadow.direction = 190;
            shadow.dist = 1;
            shadow.angle = Angle.TOP_LEFT;
            effect.shadow = shadow;

            effect.geometry = ShapeType.rect;

            Outline ln = new Outline();
            effect.outline = ln;
            ln.width = 15;
            ln.color = Color.WHITE;
            ln.cap = Outline.Cap.ROUND;

            Scene3D scene = new Scene3D();
            effect.scene3D = scene;
            Camera camera = new Camera();
            camera.presetCamera = Camera.PresetCamera.perspectiveContrastingLeftFacing;
            camera.latitude = 9;
            camera.longitude = 35;
            scene.camera = camera;
            LightRig lightRig = new LightRig();
            lightRig.rig = LightRig.Rig.soft;
            lightRig.angle = Angle.TOP;
            scene.lightRig = lightRig;

            Shape3D shape = new Shape3D();
            effect.shape3D = shape;
            shape.material = Shape3D.Material.matte;
            Bevel bevel = new Bevel();
            shape.bevelTop = bevel;
            bevel.width = 5;
            bevel.height = 4;
            shape.contourColor = new Color(192, 192, 192);
            shape.contourWidth = 1;
        }
    },
    // 25
    ReflectedBevel_Black {
        @Override public void preset(Picture pict) {
            PictureEffect effect = new PictureEffect();
            pict.effect = effect;

            Fill.SolidFill fill = new Fill.SolidFill();
            fill.color = Color.WHITE;
            effect.fill = fill;

            Reflection reflection = new Reflection();
            reflection.blur = 1;
            reflection.alpha = 72;
            reflection.size = 28;
            reflection.dist = 0.4;
            effect.reflection = reflection;

            effect.geometry = ShapeType.roundRect;
            effect.geometryAdjustValueList = Collections.singletonList(new Guide("adj", "val 4167"));

            Outline ln = new Outline();
            effect.outline = ln;
            ln.width = 6;
            ln.color = new Color(41, 41, 41);
            ln.cap = Outline.Cap.SQUARE;
            ln.joinType = Outline.JoinType.miter;
            ln.miterLimit = 800;

            Scene3D scene = new Scene3D();
            effect.scene3D = scene;
            Camera camera = new Camera();
            camera.presetCamera = Camera.PresetCamera.orthographicFront;
            scene.camera = camera;
            LightRig lightRig = new LightRig();
            lightRig.rig = LightRig.Rig.threePt;
            lightRig.angle = Angle.TOP;
            lightRig.revolution = 45;
            scene.lightRig = lightRig;

            Shape3D shape = new Shape3D();
            effect.shape3D = shape;
            Bevel bevel = new Bevel();
            shape.bevelTop = bevel;
            bevel.height = 3;
            shape.contourColor = new Color(192, 192, 192);
        }
    },
    // 26
    ReflectedBevel_White {
        @Override public void preset(Picture pict) {
            PictureEffect effect = new PictureEffect();
            pict.effect = effect;

            Fill.SolidFill fill = new Fill.SolidFill();
            fill.color = Color.WHITE;
            effect.fill = fill;

            Reflection reflection = new Reflection();
            reflection.blur = 1;
            reflection.alpha = 67;
            reflection.size = 28;
            reflection.dist = 0.4;
            effect.reflection = reflection;

            effect.geometry = ShapeType.roundRect;
            effect.geometryAdjustValueList = Collections.singletonList(new Guide("adj", "val 4167"));

            Outline ln = new Outline();
            effect.outline = ln;
            ln.width = 6;
            ln.color = new Color(234, 234, 234);
            ln.cap = Outline.Cap.SQUARE;
            ln.joinType = Outline.JoinType.miter;
            ln.miterLimit = 800;

            Scene3D scene = new Scene3D();
            effect.scene3D = scene;
            Camera camera = new Camera();
            camera.presetCamera = Camera.PresetCamera.orthographicFront;
            scene.camera = camera;
            LightRig lightRig = new LightRig();
            lightRig.rig = LightRig.Rig.threePt;
            lightRig.angle = Angle.TOP;
            lightRig.revolution = 45;
            scene.lightRig = lightRig;

            Shape3D shape = new Shape3D();
            effect.shape3D = shape;
            Bevel bevel = new Bevel();
            shape.bevelTop = bevel;
            bevel.height = 3;
            bevel.width = 6;
            shape.contourColor = new Color(192, 192, 192);
            shape.contourWidth = 0.5;
        }
    },
    // 27
    MetalRoundedRectangle {
        @Override public void preset(Picture pict) {
            PictureEffect effect = new PictureEffect();
            pict.effect = effect;
            Shadow shadow = new Shadow();
            shadow.color = Color.BLACK;
            shadow.alpha = 55;
            shadow.blur = 8;
            shadow.direction = 120;
            shadow.dist = 4;
            shadow.angle = Angle.TOP_LEFT;
            effect.shadow = shadow;

            effect.geometry = ShapeType.roundRect;
            effect.geometryAdjustValueList = Collections.singletonList(new Guide("adj", "val 11111"));

            Outline ln = new Outline();
            effect.outline = ln;
            ln.width = 15;
            ln.color = new Color(200, 198, 189);
            ln.cap = Outline.Cap.ROUND;
            ln.dash = Outline.DashPattern.solid;

            Scene3D scene = new Scene3D();
            effect.scene3D = scene;
            Camera camera = new Camera();
            camera.presetCamera = Camera.PresetCamera.perspectiveFront;
            camera.fov = 90;
            scene.camera = camera;
            LightRig lightRig = new LightRig();
            lightRig.rig = LightRig.Rig.threePt;
            lightRig.angle = Angle.TOP;
            lightRig.revolution = 320;
            scene.lightRig = lightRig;

            Shape3D shape = new Shape3D();
            effect.shape3D = shape;
            Bevel bevel = new Bevel();
            shape.bevelTop = bevel;
            bevel.width = 24;
            bevel.height = 12;
            bevel.prst = Bevel.BevelPresetType.hardEdge;
            shape.extrusionColor = Color.WHITE;
            shape.extrusionHeight = 2;
        }
    },
    // 28
    MetalOval {
        @Override public void preset(Picture pict) {
            PictureEffect effect = new PictureEffect();
            pict.effect = effect;
            Shadow shadow = new Shadow();
            shadow.color = Color.BLACK;
            shadow.blur = 10;
            shadow.angle = Angle.BOTTOM_LEFT;
            effect.shadow = shadow;

            effect.geometry = ShapeType.ellipse;

            Outline ln = new Outline();
            effect.outline = ln;
            ln.width = 15;
            ln.color = new Color(200, 198, 189);
            ln.cap = Outline.Cap.ROUND;
            ln.dash = Outline.DashPattern.solid;

            Scene3D scene = new Scene3D();
            effect.scene3D = scene;
            Camera camera = new Camera();
            camera.presetCamera = Camera.PresetCamera.perspectiveFront;
            camera.fov = 90;
            scene.camera = camera;
            LightRig lightRig = new LightRig();
            lightRig.rig = LightRig.Rig.threePt;
            lightRig.angle = Angle.TOP;
            lightRig.revolution = 320;
            scene.lightRig = lightRig;

            Shape3D shape = new Shape3D();
            effect.shape3D = shape;
            Bevel bevel = new Bevel();
            shape.bevelTop = bevel;
            bevel.width = 24;
            bevel.height = 12;
            bevel.prst = Bevel.BevelPresetType.hardEdge;
            shape.extrusionColor = Color.BLACK;
            shape.extrusionHeight = 2;
        }
    }
    ;

    /**
     * Attach preset styles to images
     *
     * @param pict {@link Picture}
     */
    public abstract void preset(Picture pict);
}

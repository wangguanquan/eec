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

import java.util.List;

import static org.ttzero.excel.drawing.Enums.ShapeType;

/**
 * Picture Effects
 *
 * @author guanquan.wang at 2023-07-25 09:24
 */
public class Effect {

    // Shadow
    public Shadow shadow, innerShadow;

    // Reflection
    public Reflection reflection;

    // Glow
    public Glow glow;

    /**
     * Soft Edges (0-100)
     *
     * Fades the existing edge of your image to the degree you specify.
     */
    public double softEdges;

    /**
     * Geometry
     *
     * Support preset geometry only
     */
    public ShapeType geometry;
    /**
     * The preset geometry can be adjusted by specifying a list of shape
     * adjustment values within a {@code AdjustValueList}
     */
    public List<Tuple2<String, String>> geometryAdjustValueList;

    // Fill
    public Fill fill;

    // Outline
    public Outline outline;

    // 3D Scene
    public Scene3D scene3D;

    // 3D Shape
    public Shape3D shape3D;

    public Shadow getShadow() {
        return shadow;
    }

    public Effect setShadow(Shadow shadow) {
        this.shadow = shadow;
        return this;
    }

    public Shadow getInnerShadow() {
        return innerShadow;
    }

    public Effect setInnerShadow(Shadow innerShadow) {
        this.innerShadow = innerShadow;
        return this;
    }

    public Reflection getReflection() {
        return reflection;
    }

    public Effect setReflection(Reflection reflection) {
        this.reflection = reflection;
        return this;
    }

    public double getSoftEdges() {
        return softEdges;
    }

    public Effect setSoftEdges(double softEdges) {
        this.softEdges = softEdges;
        return this;
    }

    public Enums.ShapeType getGeometry() {
        return geometry;
    }

    public Effect setGeometry(Enums.ShapeType geometry) {
        this.geometry = geometry;
        return this;
    }

    public List<Tuple2<String, String>> getGeometryAdjustValueList() {
        return geometryAdjustValueList;
    }

    public Effect setGeometryAdjustValueList(List<Tuple2<String, String>> geometryAdjustValueList) {
        this.geometryAdjustValueList = geometryAdjustValueList;
        return this;
    }

    public Outline getOutline() {
        return outline;
    }

    public Effect setOutline(Outline outline) {
        this.outline = outline;
        return this;
    }

    public Scene3D getScene3D() {
        return scene3D;
    }

    public Effect setScene3D(Scene3D scene3D) {
        this.scene3D = scene3D;
        return this;
    }

    public Shape3D getShape3D() {
        return shape3D;
    }

    public Effect setShape3D(Shape3D shape3D) {
        this.shape3D = shape3D;
        return this;
    }

    public Glow getGlow() {
        return glow;
    }

    public Effect setGlow(Glow glow) {
        this.glow = glow;
        return this;
    }

    public Fill getFill() {
        return fill;
    }

    public Effect setFill(Fill fill) {
        this.fill = fill;
        return this;
    }
}

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

import java.util.List;

/**
 * Picture Effects
 *
 * @author guanquan.wang at 2023-07-25 09:24
 */
public class PictureEffect {

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
    public ShapeType geometry = ShapeType.rect;
    /**
     * The preset geometry can be adjusted by specifying a list of shape
     * adjustment values within a {@code AdjustValueList}
     */
    public List<Guide> geometryAdjustValueList;

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

    public void setShadow(Shadow shadow) {
        this.shadow = shadow;
    }

    public Shadow getInnerShadow() {
        return innerShadow;
    }

    public void setInnerShadow(Shadow innerShadow) {
        this.innerShadow = innerShadow;
    }

    public Reflection getReflection() {
        return reflection;
    }

    public void setReflection(Reflection reflection) {
        this.reflection = reflection;
    }

    public double getSoftEdges() {
        return softEdges;
    }

    public void setSoftEdges(double softEdges) {
        this.softEdges = softEdges;
    }

    public ShapeType getGeometry() {
        return geometry;
    }

    public void setGeometry(ShapeType geometry) {
        this.geometry = geometry;
    }

    public List<Guide> getGeometryAdjustValueList() {
        return geometryAdjustValueList;
    }

    public void setGeometryAdjustValueList(List<Guide> geometryAdjustValueList) {
        this.geometryAdjustValueList = geometryAdjustValueList;
    }

    public Outline getOutline() {
        return outline;
    }

    public void setOutline(Outline outline) {
        this.outline = outline;
    }

    public Scene3D getScene3D() {
        return scene3D;
    }

    public void setScene3D(Scene3D scene3D) {
        this.scene3D = scene3D;
    }

    public Shape3D getShape3D() {
        return shape3D;
    }

    public void setShape3D(Shape3D shape3D) {
        this.shape3D = shape3D;
    }
}

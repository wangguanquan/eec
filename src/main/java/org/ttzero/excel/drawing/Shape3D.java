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

import static org.ttzero.excel.drawing.Enums.Material;

/**
 * @author guanquan.wang at 2023-07-26 15:21
 */
public class Shape3D {
    /**
     * Both bevel elements are empty elements with three attributes
     * which define the characteristics of the bevel.
     */
    public Bevel bevelTop, bevelBottom;
    /**
     * Specifies a preset material type, which is a combination
     * of lighting characteristics which are intended to mimic the material.
     */
    public Material material;
    /**
     * A contour is a solid filled line that surrounds the outer edge of the shape.
     */
    public Color contourColor;
    public double contourWidth;
    /**
     * An extrusion is an artificial height applied to the shape.
     */
    public Color extrusionColor;
    public double extrusionHeight;

    public Bevel getBevelTop() {
        return bevelTop;
    }

    public Shape3D setBevelTop(Bevel bevelTop) {
        this.bevelTop = bevelTop;
        return this;
    }

    public Bevel getBevelBottom() {
        return bevelBottom;
    }

    public Shape3D setBevelBottom(Bevel bevelBottom) {
        this.bevelBottom = bevelBottom;
        return this;
    }

    public Material getMaterial() {
        return material;
    }

    public Shape3D setMaterial(Material material) {
        this.material = material;
        return this;
    }

    public Color getContourColor() {
        return contourColor;
    }

    public Shape3D setContourColor(Color contourColor) {
        this.contourColor = contourColor;
        return this;
    }

    public double getContourWidth() {
        return contourWidth;
    }

    public Shape3D setContourWidth(double contourWidth) {
        this.contourWidth = contourWidth;
        return this;
    }

    public Color getExtrusionColor() {
        return extrusionColor;
    }

    public Shape3D setExtrusionColor(Color extrusionColor) {
        this.extrusionColor = extrusionColor;
        return this;
    }

    public double getExtrusionHeight() {
        return extrusionHeight;
    }

    public Shape3D setExtrusionHeight(double extrusionHeight) {
        this.extrusionHeight = extrusionHeight;
        return this;
    }
}

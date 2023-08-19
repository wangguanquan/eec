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

import static org.ttzero.excel.drawing.Enums.PresetCamera;

/**
 * @author guanquan.wang at 2023-07-26 15:21
 */
public class Camera {
    /**
     * Preset camera
     */
    public PresetCamera presetCamera;
    /**
     * A zoom can be applied to the camera position by adding a zoom
     * attribute to the {@code camera} element. It is a percentage.
     */
    public int zoom;
    /**
     * The field of view can be modified from the view set by the preset
     * camera setting by adding a fov attribute to the {@code camera} element. (0-180)
     */
    public double fov;
    /**
     * The preset placement can be altered by specifying a child {@code rot} element.
     * The {@code rot} element defines a rotation by specifying a latitude coordinate (a lat attribute),
     * a longitude coordinate (a lon attribute), and a revolution (a rev attribute) about the axis.
     */
    public double latitude, longitude, revolution;

    public PresetCamera getPresetCamera() {
        return presetCamera;
    }

    public Camera setPresetCamera(PresetCamera presetCamera) {
        this.presetCamera = presetCamera;
        return this;
    }

    public int getZoom() {
        return zoom;
    }

    public Camera setZoom(int zoom) {
        this.zoom = zoom;
        return this;
    }

    public double getFov() {
        return fov;
    }

    public Camera setFov(double fov) {
        this.fov = fov;
        return this;
    }

    public double getLatitude() {
        return latitude;
    }

    public Camera setLatitude(double latitude) {
        this.latitude = latitude;
        return this;
    }

    public double getLongitude() {
        return longitude;
    }

    public Camera setLongitude(double longitude) {
        this.longitude = longitude;
        return this;
    }

    public double getRevolution() {
        return revolution;
    }

    public Camera setRevolution(double revolution) {
        this.revolution = revolution;
        return this;
    }
}

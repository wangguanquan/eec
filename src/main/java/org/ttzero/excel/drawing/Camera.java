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

    public enum PresetCamera {
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

    public PresetCamera getPresetCamera() {
        return presetCamera;
    }

    public void setPresetCamera(PresetCamera presetCamera) {
        this.presetCamera = presetCamera;
    }

    public int getZoom() {
        return zoom;
    }

    public void setZoom(int zoom) {
        this.zoom = zoom;
    }

    public double getFov() {
        return fov;
    }

    public void setFov(double fov) {
        this.fov = fov;
    }

    public double getLatitude() {
        return latitude;
    }

    public void setLatitude(double latitude) {
        this.latitude = latitude;
    }

    public double getLongitude() {
        return longitude;
    }

    public void setLongitude(double longitude) {
        this.longitude = longitude;
    }

    public double getRevolution() {
        return revolution;
    }

    public void setRevolution(double revolution) {
        this.revolution = revolution;
    }
}

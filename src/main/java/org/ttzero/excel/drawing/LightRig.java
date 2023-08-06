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
public class LightRig {
    /**
     * Preset Properties
     */
    public Rig rig;
    /**
     * alignment
     */
    public Angle angle;
    /**
     * The preset placement can be altered by specifying a child {@code rot} element.
     * The {@code rot} element defines a rotation by specifying a latitude coordinate (a lat attribute),
     * a longitude coordinate (a lon attribute), and a revolution (a rev attribute) about the axis.
     */
    public double latitude, longitude, revolution;

    public enum Rig {
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


    public Rig getRig() {
        return rig;
    }

    public void setRig(Rig rig) {
        this.rig = rig;
    }

    public Angle getAngle() {
        return angle;
    }

    public void setAngle(Angle angle) {
        this.angle = angle;
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

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
public class Scene3D {
    /**
     * The placement and properties of the camera in the 3D scene modify the view of the scene
     */
    public Camera camera;

    /**
     * A light rig is relevant when there is a 3D bevel. The light rig defines
     * the lighting properties associated with a scene and is specified with the
     * {@code lightRig} element. It has a rig attribute which specifies a preset group of lights
     * oriented in a specific way relative to the scene.
     */
    public LightRig lightRig;

    public Camera getCamera() {
        return camera;
    }

    public Scene3D setCamera(Camera camera) {
        this.camera = camera;
        return this;
    }

    public LightRig getLightRig() {
        return lightRig;
    }

    public Scene3D setLightRig(LightRig lightRig) {
        this.lightRig = lightRig;
        return this;
    }
}

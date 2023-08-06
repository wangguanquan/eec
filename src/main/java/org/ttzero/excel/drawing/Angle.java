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
 * @author guanquan.wang at 2023-07-25 09:24
 */
public enum Angle {
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

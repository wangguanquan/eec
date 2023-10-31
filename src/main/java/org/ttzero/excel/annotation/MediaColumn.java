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


package org.ttzero.excel.annotation;

import org.ttzero.excel.drawing.PresetPictureEffect;

import java.lang.annotation.ElementType;
import java.lang.annotation.Inherited;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 指定当前列为"媒体"格式
 *
 * <p>默认情况下总是以"值"的形式导出，如果需要导出图片则必须添加{@code MediaColumn}注解，
 * 在指定属性的同时可以使用{@link #presetEffect()}设置预设图片效果</p>
 *
 * @see PresetPictureEffect
 * @author guanquan.wang at 2023-08-06 09:15
 */
@Target({ ElementType.FIELD, ElementType.METHOD })
@Retention(RetentionPolicy.RUNTIME)
@Inherited
public @interface MediaColumn {
    /**
     * 设置预设图片效果，默认无效果
     *
     * @return {@link PresetPictureEffect}
     */
    PresetPictureEffect presetEffect() default PresetPictureEffect.None;
}

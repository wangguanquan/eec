/*
 * Copyright (c) 2017-2020, guanquan.wang@yandex.com All Rights Reserved.
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


package org.ttzero.excel.entity;

import org.ttzero.excel.entity.style.Font;

/**
 * 批注，包含标题和内容且两者且少有一个不为空
 *
 * @author guanquan.wang at 2020-05-22 09:32
 */
public class Comment {
    /**
     * 标题，加粗显示
     */
    private String title;
    /**
     * 内容
     */
    private String value;
    /**
     * 指定批注框显示的宽和高
     */
    private Double width, height;
    /**
     * 指定批注字体
     */
    private Font titleFont, valueFont;

    public Comment() { }

    public Comment(String value) {
        this.value = value;
    }

    public Comment(String title, String value) {
        this(title, value, null, null);
    }

    public Comment(String value, Double width, Double height) {
        this(null, value, width, height);
    }

    public Comment(String title, String value, Double width, Double height) {
        this.title = title;
        this.value = value;
        this.width = width;
        this.height = height;
    }

    public String getTitle() {
        return title;
    }

    public Comment setTitle(String title) {
        this.title = title;
        return this;
    }

    public Comment setTitle(String title, Font titleFont) {
        this.title = title;
        this.titleFont = titleFont;
        return this;
    }

    public String getValue() {
        return value;
    }

    public Comment setValue(String value) {
        this.value = value;
        return this;
    }

    public Comment setValue(String value, Font valueFont) {
        this.value = value;
        this.valueFont = valueFont;
        return this;
    }

    public Double getWidth() {
        return width;
    }

    public Comment setWidth(Double width) {
        this.width = width;
        return this;
    }

    public Double getHeight() {
        return height;
    }

    public Comment setHeight(Double height) {
        this.height = height;
        return this;
    }

    public Font getTitleFont() {
        return titleFont;
    }

    public Comment setTitleFont(Font titleFont) {
        this.titleFont = titleFont;
        return this;
    }

    public Font getValueFont() {
        return valueFont;
    }

    public Comment setValueFont(Font valueFont) {
        this.valueFont = valueFont;
        return this;
    }
}

/*
 * Copyright (c) 2017-2021, guanquan.wang@yandex.com All Rights Reserved.
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


package org.ttzero.excel.reader;

import org.ttzero.excel.drawing.Effect;
import org.ttzero.excel.util.StringUtil;

import java.nio.file.Path;
import java.util.List;
import java.util.stream.Collectors;

/**
 * 读取Excel图片
 *
 * @author guanquan.wang at 2021-04-24 19:53
 * @see XMLDrawings
 */
public interface Drawings {
    /**
     * 列出所有工作表包含的图片
     *
     * @return 如果存在图片时返回 {@link Picture}数组, 不存在图片返回{@code null}.
     */
    List<Picture> listPictures();

    /**
     * 列出指定工作表包含的图片
     *
     * @param sheet 指定工作表
     * @return 如果存在图片时返回 {@link Picture}数组, 不存在图片返回{@code null}.
     */
    default List<Picture> listPictures(final Sheet sheet) {
        List<Picture> pictures = listPictures();
        return pictures == null ? null :
            pictures.stream().filter(p -> p.getSheet().getId() == sheet.getId()).collect(Collectors.toList());
    }

    class Picture {
        /**
         * 图片所在的工作表 {@link Sheet}
         */
        public Sheet sheet;
        /**
         * 图片在工作表中的位置，记录图片左上角和右下角的行列坐标
         */
        public Dimension dimension;
        /**
         * 图片的临时路径
         */
        public Path localPath;
        /**
         * 如果是网络图片，则此属性保留网络图片的原始链接
         */
        public String srcUrl;
        /**
         * 是否为背景图片，水印图片
         */
        public boolean background;
        /**
         * 0: Move and size with cells
         * 1: Move but don't size with cells
         * 2: Don't move or size with cells
         */
        public int property;
        /**
         * Revolve -360 ~ 360
         */
        public int revolve;
        /**
         * Padding top | right | bottom | left
         */
        public short[] padding;
        /**
         * Picture effects
         */
        public Effect effect;

        public Sheet getSheet() {
            return sheet;
        }

        public Dimension getDimension() {
            return dimension;
        }

        public Path getLocalPath() {
            return localPath;
        }

        public String getSrcUrl() {
            return srcUrl;
        }

        public boolean isBackground() {
            return background;
        }

        public int getProperty() {
            return property;
        }

        public int getRevolve() {
            return revolve;
        }

        public short[] getPadding() {
            return padding;
        }

        public Effect getEffect() {
            return effect;
        }

        @Override
        public String toString() {
            return background ? "Background picture [" + localPath + "] in worksheet " + sheet.getName() + (StringUtil.isNotEmpty(srcUrl) ? " from internet url " + srcUrl : "")
                    : "Picture [" + localPath + "] in worksheet " + sheet.getName() + " at " + dimension + (StringUtil.isNotEmpty(srcUrl) ? " from internet url " + srcUrl : "");
        }
    }
}

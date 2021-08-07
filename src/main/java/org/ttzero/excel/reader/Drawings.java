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

import org.ttzero.excel.util.StringUtil;

import java.nio.file.Path;
import java.util.List;
import java.util.stream.Collectors;

/**
 * Drawings resources
 *
 * @author guanquan.wang at 2021-04-24 19:53
 */
public interface Drawings {
    /**
     * List all picture in excel
     *
     * @return list of {@link Picture}, or null if not exists.
     */
    List<Picture> listPictures();

    /**
     * List all picture in specify worksheet
     *
     * @param sheet Specifies witch {@code Worksheet} to get the picture from
     * @return list of {@link Picture}, or null if not exists.
     */
    default List<Picture> listPictures(final Sheet sheet) {
        List<Picture> pictures = listPictures();
        return pictures == null ? null :
            pictures.stream().filter(p -> p.sheet.getId() == sheet.getId()).collect(Collectors.toList());
    }

    class Picture {
        /**
         * Specify the {@link Sheet} which contains the picture
         */
        Sheet sheet;
        /**
         * Dimension of picture
         */
        Dimension dimension;
        /**
         * The local temporary path
         */
        Path localPath;
        /**
         * If it is an online picture, the {@code srcUrl} is the source address of the picture
         */
        String srcUrl;
        /**
         * Is background picture
         */
        boolean background;

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

        @Override
        public String toString() {
            return background ? "Background picture [" + localPath + "] in worksheet " + sheet.getName() + (StringUtil.isNotEmpty(srcUrl) ? " from internet url " + srcUrl : "")
                    : "Picture [" + localPath + "] in worksheet " + sheet.getName() + " at " + dimension + (StringUtil.isNotEmpty(srcUrl) ? " from internet url " + srcUrl : "");
        }
    }
}

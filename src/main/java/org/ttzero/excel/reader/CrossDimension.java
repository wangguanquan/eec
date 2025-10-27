/*
 * Copyright (c) 2017-2025, guanquan.wang@hotmail.com All Rights Reserved.
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

import java.util.Objects;

/**
 * 跨工作表维度
 *
 * @author wangguanquan3 on 2025-09-19
 */
public class CrossDimension {
    /**
     * 工作表名
     */
    public String sheetName;
    /**
     * 所处维度
     */
    public Dimension dimension;

    public CrossDimension(Dimension dimension) {
        this.dimension = dimension;
    }

    public CrossDimension(String sheetName, Dimension dimension) {
        this.sheetName = sheetName;
        this.dimension = dimension;
    }

    public static CrossDimension of(String referer) {
        int i = referer.indexOf('!');
        String otherSheetName = null;
        Dimension dimension = null;
        if (i > 0 && i < referer.length() - 2) {
            otherSheetName = referer.substring(0, i);
            referer = referer.substring(i + 1);
        }
        try {
            dimension = Dimension.of(referer.replace("$", ""));
        } catch (Exception ex) {
            // Ignore
        }
        if (dimension == null) throw new IllegalArgumentException("Invalid dimension:" + referer);
        return new CrossDimension(otherSheetName, dimension);
    }

    /**
     * 判断是否跨工作表
     *
     * @return {@code true}跨工作表
     */
    public boolean isCrossSheet() {
        return StringUtil.isNotEmpty(sheetName);
    }

    @Override
    public int hashCode() {
        int h = dimension.hashCode();
        return isCrossSheet() ? h ^ sheetName.hashCode() : h;
    }

    @Override
    public boolean equals(Object o) {
        boolean r = this == o;
        if (!r && o instanceof CrossDimension) {
            CrossDimension other = (CrossDimension) o;
            r = Objects.equals(other.sheetName, sheetName) && other.dimension.equals(dimension);
        }
        return r;
    }

    @Override
    public String toString() {
        return StringUtil.isNotEmpty(sheetName) ? (sheetName + "!" + dimension.toReferer()) : dimension.toString();
    }
}

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

package org.ttzero.excel.entity.e7;

import org.ttzero.excel.entity.Column;
import org.ttzero.excel.entity.style.Fill;
import org.ttzero.excel.entity.style.Styles;

/**
 * @author guanquan.wang at 2023-02-24 11:12
 */
public class XMLZebraLineCellValueAndStyle extends XMLCellValueAndStyle {

    /**
     * The zebra-line fill style value
     */
    protected int zebraFillStyle = -1;
    /**
     * The zebra-line fill style
     */
    protected Fill zebraFill;

    public XMLZebraLineCellValueAndStyle(int zebraFillStyle) {
        this.zebraFillStyle = zebraFillStyle;
    }

    public XMLZebraLineCellValueAndStyle(Fill zebraFill) {
        this.zebraFill = zebraFill;
    }

    /**
     * Returns the cell style index
     *
     * @param hc    the header column
     * @param o     the cell value
     * @param style the default style
     * @return the style index in xf
     */
    @Override
    protected int getStyleIndex(int rows, Column hc, Object o, int style) {
        if (zebraFillStyle == -1 && zebraFill != null)
            zebraFillStyle = hc.styles.addFill(zebraFill);
        // Interlaced discoloration
        if (isOdd(rows) && !Styles.hasFill(style)) style |= zebraFillStyle;
        return super.getStyleIndex(rows, hc, o, style);
    }


    /**
     * Check the odd rows
     *
     * @return true if odd rows
     */
    static boolean isOdd(int rows) {
        return (rows & 1) == 1;
    }

}

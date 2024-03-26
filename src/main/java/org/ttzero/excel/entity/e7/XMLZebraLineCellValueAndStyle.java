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
import org.ttzero.excel.entity.ICellValueAndStyle;
import org.ttzero.excel.entity.Row;
import org.ttzero.excel.entity.style.Fill;
import org.ttzero.excel.entity.style.Styles;

/**
 * 斑马线样式，默认情况下从数据行开始计算，每隔一行添加指定填充色，默认填充色为 {@code #E9EAEC}，
 * 如果该单元格已有填充样式则保持原填充样式
 *
 * @author guanquan.wang at 2023-02-24 11:12
 */
public class XMLZebraLineCellValueAndStyle implements ICellValueAndStyle {

    /**
     * 斑马线填充样值
     */
    protected int zebraFillStyle = -1;
    /**
     * 斑马线填充样式
     */
    protected Fill zebraFill;

    public XMLZebraLineCellValueAndStyle(int zebraFillStyle) {
        this.zebraFillStyle = zebraFillStyle;
    }

    public XMLZebraLineCellValueAndStyle(Fill zebraFill) {
        this.zebraFill = zebraFill;
    }

    /**
     * 获取单元格样式值，先通过{@code Column}获取基础样式并在偶数行添加斑马线填充，
     * 如果有动态样式转换则将基础样式做为参数进行二次制作
     *
     * @param row 行信息
     * @param hc  当前列的表头
     * @param o   单元格的值
     * @return 样式值
     */
    @Override
    public int getStyleIndex(Row row, Column hc, Object o) {
        if (zebraFillStyle == -1 && zebraFill != null)
            zebraFillStyle = hc.styles.addFill(zebraFill);
        // Default style
        int style = hc.getCellStyle();
        // 偶数行且无特殊填充样式时添加斑马线填充
        if (isOdd(row.getIndex()) && !Styles.hasFill(style)) style |= zebraFillStyle;
        // 处理动态样式
        if (hc.styleProcessor != null) {
            style = hc.styleProcessor.build(o, style, hc.styles);
        }
        return hc.styles.of(style);
    }


    /**
     * 检查是否需要添加斑马线样式
     *
     * @param rows 数据行的行号（zero base)
     * @return 数据行的偶数行返回 {@code true}
     */
    public static boolean isOdd(int rows) {
        return (rows & 1) == 1;
    }

}

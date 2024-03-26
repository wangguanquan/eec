/*
 * Copyright (c) 2017-2019, guanquan.wang@yandex.com All Rights Reserved.
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

import org.ttzero.excel.entity.ICellValueAndStyle;

/**
 * @author guanquan.wang at 2019-09-25 11:25
 * @deprecated 即将删除，直接实现 {@link ICellValueAndStyle}即可
 */
@Deprecated
public class XMLCellValueAndStyle implements ICellValueAndStyle {

//    /**
//     * Int value conversion to others
//     *
//     * @param row  the row number
//     * @param cell the cell
//     * @param o    the cell value
//     * @param hc   the header column
//     */
//    @Override
//    public void conversion(Row row, Cell cell, Object o, Column hc) {
//        Object e = hc.getConversion().conversion(o);
//        if (e != null) {
//            Class<?> clazz = e.getClass();
//            setCellValue(row, cell, e, hc, clazz, false);
//            // FIXME 转转换后是否根据转换后的类型重新设置对齐不明确，暂时不重置
////            int style = hc.getCellStyle(clazz);
//            cell.xf = getStyleIndex(row, hc, o, hc.getCellStyle());
//        } else {
//            cell.blank();
//            cell.xf = getStyleIndex(row, hc, null, hc.getCellStyle());
//        }
//    }

}

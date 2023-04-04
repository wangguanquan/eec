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

import org.ttzero.excel.entity.Column;
import org.ttzero.excel.entity.ICellValueAndStyle;
import org.ttzero.excel.manager.Const;
import org.ttzero.excel.processor.StyleProcessor;
import org.ttzero.excel.reader.Cell;

import static org.ttzero.excel.entity.IWorksheetWriter.isChar;
import static org.ttzero.excel.entity.IWorksheetWriter.isInt;
import static org.ttzero.excel.entity.IWorksheetWriter.isShort;

/**
 * @author guanquan.wang at 2019-09-25 11:25
 */
public class XMLCellValueAndStyle implements ICellValueAndStyle {

    /**
     * Int value conversion to others
     *
     * @param cell the cell
     * @param o    the cell value
     * @param hc   the header column
     */
    @Override
    public void conversion(int row, Cell cell, Object o, Column hc) {
        Object e = hc.processor.conversion(o);
        if (e != null) {
            Class<?> clazz = e.getClass();
            if (isInt(clazz)) {
                if (isChar(clazz)) {
                    cell.setCv((Character) e);
                } else if (isShort(clazz)) {
                    cell.setNv((Short) e);
                } else {
                    cell.setNv((Integer) e);
                }
                cell.xf = getStyleIndex(row, hc, e);
            } else {
                setCellValue(row, cell, e, hc, clazz, false);
                // FIXME Here will override the style set by the user
                int style = hc.getCellStyle(clazz);
                cell.xf = getStyleIndex(row, hc, o, style);
            }
        } else {
            cell.blank();
            cell.xf = getStyleIndex(row, hc, null);
        }
    }

    /**
     * Setting cell value and cell styles
     *
     * @param cell the cell
     * @param e    the cell value
     * @param hc   the header column
     */
    @Deprecated
    @Override
    public void reset(int row, Cell cell, Object e, Column hc) {
        setCellValue(row, cell, e, hc, hc.getClazz(), hc.processor != null);
        if (hc.processor == null) {
            cell.xf = getStyleIndex(row, hc, e);
        }
    }

    /**
     * Returns the cell style index
     *
     * @param hc    the header column
     * @param o     the cell value
     * @param style the default style
     * @return the style index in xf
     */
    @Deprecated
    protected int getStyleIndex(int rows, Column hc, Object o, int style) {
        if (hc.styleProcessor != null) {
            style = hc.styleProcessor.build(o, style, hc.styles);
        }
        return hc.styles.of(style);
    }

    /**
     * Returns the cell style index
     *
     * @param hc the header column
     * @param o  the cell value
     * @return the style index in xf
     */
    @Deprecated
    @Override
    public int getStyleIndex(int rows, Column hc, Object o) {
        int style = hc.getCellStyle();
        return getStyleIndex(rows, hc, o, style);
    }

    /**
     * Returns the worksheet name
     *
     * @return name of worksheet
     */
    @Override
    public String getFileSuffix() {
        return Const.Suffix.XML;
    }
}

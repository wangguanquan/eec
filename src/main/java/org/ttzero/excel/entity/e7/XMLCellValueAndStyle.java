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
import org.ttzero.excel.entity.style.Styles;
import org.ttzero.excel.manager.Const;
import org.ttzero.excel.reader.Cell;

import static org.ttzero.excel.entity.IWorksheetWriter.isChar;
import static org.ttzero.excel.entity.IWorksheetWriter.isInt;
import static org.ttzero.excel.entity.IWorksheetWriter.isShort;

/**
 * @author guanquan.wang at 2019-09-25 11:25
 */
public class XMLCellValueAndStyle implements ICellValueAndStyle {
    /**
     * Automatic interlacing color
     */
    private final int autoOdd;
    /**
     * Odd row's background color
     */
    private final int oddFill;

    public XMLCellValueAndStyle(int autoOdd, int oddFill) {
        this.autoOdd = autoOdd;
        this.oddFill = oddFill;
    }

    /**
     * Int value conversion to others
     *
     * @param cell the cell
     * @param n    the cell value
     * @param hc   the header column
     */
    @Override
    public void conversion(int row, Cell cell, int n, Column hc) {
        Object e = hc.processor.conversion(n);
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
                setCellValue(row, cell, e, hc, clazz);
                int style = hc.getCellStyle(clazz);
                cell.xf = getStyleIndex(row, hc, n, style);
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
    @Override
    public void reset(int row, Cell cell, Object e, Column hc) {
        setCellValue(row, cell, e, hc, hc.getClazz());
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
    private int getStyleIndex(int rows, Column hc, Object o, int style) {
        // Interlaced discoloration
        if (autoOdd == 0 && isOdd(rows) && !Styles.hasFill(style)) {
            style |= oddFill;
        }
        int styleIndex = hc.styles.of(style);
        if (hc.styleProcessor != null) {
            style = hc.styleProcessor.build(o, style, hc.styles);
            styleIndex = hc.styles.of(style);
        }
        return styleIndex;
    }

    /**
     * Returns the cell style index
     *
     * @param hc the header column
     * @param o  the cell value
     * @return the style index in xf
     */
    @Override
    public int getStyleIndex(int rows, Column hc, Object o) {
        int style = hc.getCellStyle();
        return getStyleIndex(rows, hc, o, style);
    }

    /**
     * Check the odd rows
     *
     * @return true if odd rows
     */
    private boolean isOdd(int rows) {
        return (rows & 1) == 1;
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

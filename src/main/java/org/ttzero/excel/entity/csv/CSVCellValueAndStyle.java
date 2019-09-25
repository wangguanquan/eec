/*
 * Copyright (c) 2019, guanquan.wang@yandex.com All Rights Reserved.
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

package org.ttzero.excel.entity.csv;

import org.ttzero.excel.entity.ICellValueAndStyle;
import org.ttzero.excel.entity.Sheet;
import org.ttzero.excel.manager.Const;
import org.ttzero.excel.reader.Cell;

/**
 * Create by guanquan.wang at 2019-09-25 11:46
 */
public class CSVCellValueAndStyle implements ICellValueAndStyle {
    /**
     * Setting cell value and cell styles
     *
     * @param cell the cell
     * @param e    the cell value
     * @param hc   the header column
     */
    @Override
    public void reset(int row, Cell cell, Object e, Sheet.Column hc) {
        setCellValue(row, cell, e, hc, hc.getClazz());
    }

    /**
     * Returns the worksheet name
     *
     * @return name of worksheet
     */
    @Override
    public String getFileSuffix() {
        return Const.Suffix.CSV;
    }
}

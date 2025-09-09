/*
 * Copyright (c) 2017-2022, guanquan.wang@hotmail.com All Rights Reserved.
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

import java.util.List;

/**
 * @author guanquan.wang at 2022-07-04 11:56
 */
public class GridFactory {
    private GridFactory() { }
    public static Grid create(List<Dimension> mergeCells) {
        Dimension dim = mergeCells.get(0);
        int fr = dim.firstRow, lr = dim.lastRow;
        short fc = dim.firstColumn, lc = dim.lastColumn;
        int n = (lr - fr + 1) * (lc - fc + 1);
        for (int j = 1, len = mergeCells.size(); j < len; j++) {
            dim = mergeCells.get(j);
            n += (dim.lastRow - dim.firstRow + 1) * (dim.lastColumn - dim.firstColumn + 1);
            if (fr > dim.firstRow)    fr = dim.firstRow;
            if (lr < dim.lastRow)     lr = dim.lastRow;
            if (fc > dim.firstColumn) fc = dim.firstColumn;
            if (lc < dim.lastColumn)  lc = dim.lastColumn;
        }

        Dimension range = new Dimension(fr, fc, lr, lc);
        int r = lr - fr + 1, c = lc - fc + 1;
        n = r * c;

        Grid grid = c <= 64 && r < 1 << 15 ? new Grid.FastGrid(range)
            : n > 1 << 17 ? new Grid.FractureGrid(range) : new Grid.IndexGrid(range, n);

        for (Dimension d : mergeCells) grid.mark(d);
        return grid;
    }
}

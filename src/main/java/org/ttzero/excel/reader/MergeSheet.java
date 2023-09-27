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


package org.ttzero.excel.reader;

import java.util.List;

/**
 * Copy values when reading merged cells.
 * <p>
 * By default, the values of the merged cells are only
 * stored in the first Cell, and other cells have no values.
 * Call this method to copy the value to other cells in the merge.
 * <blockquote><pre>
 * |---------|     |---------|     |---------|
 * |         |     |  1 |    |     |  1 |  1 |
 * |    1    |  =&gt; |----|----|  =&gt; |----|----|
 * |         |     |    |    |     |  1 |  1 |
 * |---------|     |---------|     |---------|
 * Merged(A1:B2)     Default           Copy
 *                  Value in A1
 *                  others are
 *                  `null`
 * </pre></blockquote>
 * @author guanquan.wang at 2022-08-10 11:36
 */
public interface MergeSheet extends Sheet {
    /**
     * Returns CellMerged info
     *
     * @return merged {@link Grid}
     */
    Grid getMergeGrid();

    /**
     * Returns all merged cells
     *
     * @return If no merged cells are returned, Null is returned
     */
    List<Dimension> getMergeCells();
}

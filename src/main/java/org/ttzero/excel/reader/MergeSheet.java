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
 * 支持复制合并单元格的工作表，可以通过{@link #asMergeSheet}将普通工作表转为{@code MergeSheet}
 *
 * <p>通常合并单元格的值保存在左上角第一个单元格中其余单元格的值为{@code null}，如果要读取这类合并单元的值就需要特殊处理，
 * 如果将工作表转为{@code MergeSheet}就可以直接获取合并范围内的所有单元格的值，每个值均为首个单元格的值。
 * 使用{@code MergeSheet}会有一定的性能损耗，数据量少时可忽略这种损耗</p>
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
 *
 * @author guanquan.wang at 2022-08-10 11:36
 */
public interface MergeSheet extends Sheet {
    /**
     * 获取抽象的合并表格，通过表格快速判断某个坐标是否为合并单元格的一部分
     *
     * @return merged {@link Grid}
     */
    Grid getMergeGrid();

    /**
     * 获取所有合并单元格的合并范围
     *
     * @return 如果存在合并单元格则返回所有合并单元格的范围，否则返回{@code null}
     */
    List<Dimension> getMergeCells();
}

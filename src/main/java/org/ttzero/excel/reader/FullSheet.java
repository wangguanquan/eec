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


package org.ttzero.excel.reader;

import org.ttzero.excel.entity.Comment;
import org.ttzero.excel.entity.Panes;

import java.util.List;
import java.util.Map;

/**
 * 全属性工作表，与普通工作表不同除了值以外{@code FullSheet}将会额外读取行高和列宽以及单元格公式，
 * 全属性工作表虽然继承{@code MergeSheet}但并不会主动在合并单元格复制值，如果需要复制值需要明确调用{@link #copyOnMerged}方法
 *
 * @author guanquan.wang at 2023-12-02 15:31
 */
public interface FullSheet extends MergeSheet, CalcSheet {
    /**
     * 复制合并单元格的值
     *
     * <p>通常合并单元格的值保存在左上角第一个单元格中其余单元格的值为{@code null}，如果要读取这类合并单元的值就需要特殊处理，
     * 使用{@code copyOnMerged}方法就可以直接获取合并范围内的所有单元格的值，每个值均为首个单元格的值。
     * 本方法会有一定的性能损耗，数据量少时可忽略这种损耗</p>
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
     * @return 本工作表
     */
    FullSheet copyOnMerged();
    /**
     * 获取冻结信息
     *
     * @return 冻结的行列信息，无冻结时返回{@code null}
     */
    Panes getFreezePanes();

    /**
     * 获取列宽相Cols属性
     *
     * @return Cols列表，可以为{@code null}
     */
    List<Col> getCols();

    /**
     * 获取筛选区域，自动筛选的配置放在工作表最后，所以获取此值耗时较长
     *
     * @return 范围值，没有筛选时返回{@code null}
     */
    Dimension getFilter();

    /**
     * 获取预置列宽，该列宽优先级最低可以被{@link #getCols()}里的列宽覆盖，有效范围以外的列会展示此宽度
     *
     * @return 预置列宽，{@code -1}表示未知
     */
    double getDefaultColWidth();
    /**
     * 获取预置行高，暂不知道何种场景下此行高生交
     *
     * @return 预置行高，{@code -1}表示未知
     */
    double getDefaultRowHeight();

    /**
     * 工作表是否显示网络线
     *
     * @return true: 显示网络线
     */
    boolean isShowGridLines();

    /**
     * 获取工作表缩放比例，取值{@code 10-400}
     *
     * @return 百分比整数化，{@code null}表示未设置
     */
    Integer getZoomScale();

    /**
     * 获取批注
     *
     * @return key: 行列值 {@code col & 0x7FFF | ((long) row) << 16}, value: 批注
     */
    Map<Long, Comment> getComments();
}

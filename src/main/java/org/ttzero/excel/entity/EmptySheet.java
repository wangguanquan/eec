/*
 * Copyright (c) 2017-2018, guanquan.wang@hotmail.com All Rights Reserved.
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

package org.ttzero.excel.entity;

import java.util.Arrays;
import java.util.List;

/**
 * 空工作表，可用于占位，如果指定表头则会输出表头
 *
 * @author guanquan.wang at 2018-01-29 16:05
 */
public class EmptySheet extends Sheet {

    /**
     * 实例化工作表，未指定工作表名称时默认以{@code 'Sheet'+id}命名
     */
    public EmptySheet() {
        super();
    }

    /**
     * 实例化工作表并指定工作表名称
     *
     * @param name 工作表名称
     */
    public EmptySheet(String name) {
        super(name);
    }

    /**
     * 实例化工作表并指定表头信息
     *
     * @param columns 表头信息
     */
    public EmptySheet(Column... columns) {
        super(columns);
    }

    /**
     * 实例化工作表并指定工作表名称和表头信息
     *
     * @param name    工作表名称
     * @param columns 表头信息
     */
    public EmptySheet(String name, final Column... columns) {
        super(name, columns);
    }

    /**
     * 实例化工作表并指定工作表名称，水印和表头信息
     *
     * @param name      工作表名称
     * @param watermark 水印
     * @param columns   表头信息
     * @deprecated 使用场景极少，后续版本将删除
     */
    @Deprecated
    public EmptySheet(String name, Watermark watermark, final Column... columns) {
        super(name, watermark, columns);
    }

    /**
     * Reset the row-block data
     */
    @Override
    protected void resetBlockData() { }

    /**
     * Returns total rows in this worksheet
     *
     * @return 0
     */
    @Override
    public int size() {
        return 0;
    }

    /**
     * 设置表头信息，与Columns不同的是本方法只设置表头值并不带任何其它属性，可以看为{@link #setColumns(List)}的简化方法
     *
     * @param header 表头信息列表
     * @return 当前对象，支持链式调用
     */
    public EmptySheet setHeader(List<String> header) {
        Column[] columns;
        if (header == null || header.isEmpty()) columns = new Column[0];
        else {
            columns = new Column[header.size()];
            for (int i = 0, len = header.size(); i < len; columns[i] = new Column(header.get(i++)).setCellStyle(0));
        }
        super.setColumns(columns);
        return this;
    }

    /**
     * 设置表头信息，与Columns不同的是本方法只设置表头值并不带任何其它属性，可以看为{@link #setColumns(Column...)}的简化方法
     *
     * @param header 表头信息列表
     * @return 当前对象，支持链式调用
     */
    public EmptySheet setHeader(String ... header) {
        return setHeader(Arrays.asList(header));
    }
}

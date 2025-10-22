/*
 * Copyright (c) 2017-2025, guanquan.wang@hotmail.com All Rights Reserved.
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

import java.io.IOException;
import java.util.List;

/**
 * 标记支持推送模式的Worksheet
 *
 * @author guanquan.wang at 2025-09-10 10:30
 */
public interface IPushModelSheet<T> {
    /**
     * 写数据，{@code PUSH}模式下数据将直接被写到文件中
     *
     * @param data 待写入的数据
     * @return 已写入的总数据行，不包含表头
     * @throws IOException if I/O error occur.
     */
    int writeData(List<T> data) throws IOException;
}

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


package org.ttzero.excel.entity;

import java.io.Closeable;
import java.io.IOException;

/**
 * 多媒体类型输出协议（目前只支持图片）
 *
 * @author guanquan.wang at 2023-03-07 09:09
 */
public interface IDrawingsWriter extends Closeable, Storable {
    /**
     * 添加图片
     *
     * @param picture 图片信息{@link Picture}
     * @throws IOException if I/O error occur.
     */
    void drawing(Picture picture) throws IOException;

    /**
     * 异步添加图片
     *
     * @param picture 图片信息{@link Picture}
     * @throws IOException if I/O error occur.
     */
    void asyncDrawing(Picture picture) throws IOException;

    /**
     * 通知图片忆准备好，与{@link #asyncDrawing}搭配使用
     *
     * @param picture 已完成的图片{@link Picture}
     */
    void complete(Picture picture);
}

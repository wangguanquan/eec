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

package cn.ttzero.excel.processor;

import java.nio.file.Path;

/**
 * Excel定操作完成后可以做后续操作
 * Created by guanquan.wang on 2018/6/13.
 */
@FunctionalInterface
public interface DownProcessor {
    /**
     * 执行此方法
     * @param path excel临时位置
     */
    void exec(Path path);
}

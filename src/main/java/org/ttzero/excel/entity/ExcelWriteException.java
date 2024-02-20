/*
 * Copyright (c) 2017, guanquan.wang@yandex.com All Rights Reserved.
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

/**
 * Excel导出异常，收集导出所需数据或者参数与预期不符时抛出此异常，通常它不用于写异常，写磁盘异常统一使用{@code IOException}，
 * 也就是说{@code ExcelWriteException}是在数据收集阶段它抛出的时机要早于{@code IOException}
 *
 * @author guanquan.wang on 2017/10/19.
 */
public class ExcelWriteException extends RuntimeException {

    private static final long serialVersionUID = 1L;

    public ExcelWriteException() {
        super();
    }

    public ExcelWriteException(String s) {
        super(s);
    }

    public ExcelWriteException(String message, Throwable cause) {
        super(message, cause);
    }

    public ExcelWriteException(Throwable cause) {
        super(cause);
    }
}

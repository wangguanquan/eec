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
 * The maximum number of columns in the xlsx file is 16_384.
 * If this number is exceeded, this exception will occur.
 *
 * @author guanquan.wang on 2017/10/19.
 */
public class TooManyColumnsException extends ExcelWriteException {

    private static final long serialVersionUID = 1L;

    public TooManyColumnsException() {
        super();
    }

    public TooManyColumnsException(int n, int m) {
        super(n + " out of Total number of columns on a worksheet " + m);
    }

    public TooManyColumnsException(String s) {
        super(s);
    }
}

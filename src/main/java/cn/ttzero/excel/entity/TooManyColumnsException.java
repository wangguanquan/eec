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

package cn.ttzero.excel.entity;

import cn.ttzero.excel.manager.Const;

/**
 * xlsx文件最大列数为16_384，如果超出这个数将抛出此异常
 * Created by guanquan.wang on 2017/10/19.
 */
public class TooManyColumnsException extends ExportException {

	private static final long serialVersionUID = 1L;

	public TooManyColumnsException() {
        super();
    }

    public TooManyColumnsException(int n) {
        super(n + " out of Total number of columns on a worksheet " + Const.Limit.MAX_COLUMNS_ON_SHEET);
    }

    public TooManyColumnsException(String s) {
        super(s);
    }
}

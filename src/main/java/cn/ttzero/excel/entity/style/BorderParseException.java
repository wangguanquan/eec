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

package cn.ttzero.excel.entity.style;

/**
 * Created by guanquan.wang at 2018-02-08 13:53
 */
public class BorderParseException extends RuntimeException {

    private static final long serialVersionUID = 1L;

    public BorderParseException(String message) {
        super(message);
    }

    public BorderParseException(String message, Throwable cause) {
        super(message, cause);
    }

    public BorderParseException(Throwable cause) {
        super(cause);
    }

    protected BorderParseException(String message, Throwable cause,
                                   boolean enableSuppression,
                                   boolean writableStackTrace) {
        super(message, cause, enableSuppression, writableStackTrace);
    }
}

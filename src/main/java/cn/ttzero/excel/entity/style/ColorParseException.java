/*
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
 * Created by guanquan.wang at 2018-02-09 14:25
 */
public class ColorParseException extends RuntimeException {
    public ColorParseException(String message) {
    super(message);
}

    public ColorParseException(String message, Throwable cause) {
        super(message, cause);
    }

    public ColorParseException(Throwable cause) {
        super(cause);
    }

    protected ColorParseException(String message, Throwable cause,
                                   boolean enableSuppression,
                                   boolean writableStackTrace) {
        super(message, cause, enableSuppression, writableStackTrace);
    }
}

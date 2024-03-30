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


package org.ttzero.excel.validation;

import org.ttzero.excel.util.DateUtil;
import org.ttzero.excel.util.StringUtil;

import java.text.ParsePosition;
import java.util.Date;

/**
 * 日期验证，限定起始和结束时间范围
 *
 * @author guanquan.wang at 2022-08-17 20:05:42
 */
public class DateValidation extends Tuple2Validation<Integer, Integer> {
    public DateValidation between(Date from, Date to) {
        if (from != null) v1 = DateUtil.toDateValue(from);
        if (to != null) v2 = DateUtil.toDateValue(to);
        return this;
    }

    /**
     * @param from time in format "yyyy-MM-dd"
     * @param to   time in format "yyyy-MM-dd"
     * @return DateValidation
     */
    public DateValidation between(String from, String to) {
        if (StringUtil.isNotEmpty(from))
            v1 = DateUtil.toDateValue(DateUtil.dateFormat.get().parse(from.substring(0, Math.min(from.length(), 10)), new ParsePosition(0)));
        if (StringUtil.isNotEmpty(to))
            v2 = DateUtil.toDateValue(DateUtil.dateFormat.get().parse(to.substring(0, Math.min(to.length(), 10)), new ParsePosition(0)));
        return this;
    }

    @Override
    public String getType() {
        return "date";
    }

}

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


package org.ttzero.excel.validation;

import java.math.BigDecimal;

/**
 * 小数验证
 *
 * @author guanquan.wang on 2025-10-10
 */
public class DecimalValidation extends RangeValidation<BigDecimal> {
    @Override
    public String getType() {
        return "decimal";
    }

    @Override
    protected BigDecimal parseTxtValue(String txt) {
        return new BigDecimal(txt);
    }
}

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

package cn.ttzero.excel.entity.e7;

import cn.ttzero.excel.entity.AbstractTemplate;
import cn.ttzero.excel.entity.Workbook;

import java.nio.file.Path;

/**
 * Created by guanquan.wang at 2018-02-23 17:19
 */
public class SimpleTemplate extends AbstractTemplate {

    public SimpleTemplate(Path zipPath, Workbook wb) {
        super(zipPath, wb);
    }

    @Override
    protected boolean isPlaceholder(String txt) {
        int len = txt.length();
        return len > 3 &&  txt.charAt(0) == '$' && txt.charAt(1) == '{' && txt.charAt(len - 1) == '}';
    }

    private String getKey(String txt) {
        return txt.substring(2, txt.length() - 1).trim();
    }

    @Override
    protected String getValue(String txt) {
        if (map == null) return txt;
        String value, key = getKey(txt);
        if (map.containsKey(key)) {
            value = map.get(key);
        } else value = txt;

        return value;
    }

}

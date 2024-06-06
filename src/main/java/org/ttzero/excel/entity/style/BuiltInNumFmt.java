/*
 * Copyright (c) 2017-2018, guanquan.wang@yandex.com All Rights Reserved.
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

package org.ttzero.excel.entity.style;

import org.ttzero.excel.util.StringUtil;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;

/**
 * Load the Built-In Number Format
 *
 * @author guanquan.wang at 2018-02-12 10:11
 */
public final class BuiltInNumFmt {
    private static final NumFmt[] buildInNumFmts;

    static {
        InputStream is = BuiltInNumFmt.class.getClassLoader().getResourceAsStream("numFmt");
        if (is != null) {
            List<NumFmt> list = new ArrayList<>();
            int maxLen = 0;
            try (BufferedReader br = new BufferedReader(new InputStreamReader(is ,StandardCharsets.UTF_8))) {
                String v;
                String lang = Locale.getDefault().toLanguageTag();
                boolean damaged = false;
                while ((v = br.readLine()) != null) {
                    if (StringUtil.isEmpty(v)) continue;
                    v = v.trim();
                    if (v.charAt(0) == '[') {
                        int end = v.indexOf(']');
                        // 匹配方言
                        damaged = end == -1 || end == 1 || !lang.equalsIgnoreCase(v.substring(0, end).trim());
                    } else {
                        if (damaged) continue;
                        int index = v.indexOf('=');
                        if (index <= 0) continue;
                        String v1 = v.substring(0, index).trim()
                            , v2 = v.substring(index + 1).trim();
                        // check id and check code
                        int id;
                        try {
                            id = Integer.parseInt(v1);
                        } catch (NumberFormatException e) {
                            continue; // Id error.
                        }
                        if (v2.charAt(0) != '\'' || v2.charAt(v2.length() - 1) != '\'') {
                            continue; // Code error.
                        }
                        NumFmt fmt = new NumFmt();
                        fmt.id = id;
                        fmt.code = v2.substring(1, v2.length() - 1);
                        list.add(fmt);

                        if (fmt.id > maxLen) maxLen = fmt.id;
                    }
                }
            } catch (IOException e) {
                // Empty
            }

            buildInNumFmts = new NumFmt[maxLen + 1];
            // 检查ID重复ID，复制NumFmt到buildInNumFmts
            for (NumFmt fmt : list) buildInNumFmts[fmt.id] = fmt;
        } else {
            buildInNumFmts = new NumFmt[0];
        }
    }

    public static int indexOf(String code) {
        NumFmt v = get(code);
        return v != null ? v.id : -1;
    }

    /**
     * 按名称查找内置格式化对象，注意此方法不兼容方言，如需支持zh-CN以外的方言需在numFmt文件中添加对应配置
     *
     * @param code 格式化Code
     * @return 对应 {@link NumFmt}，如果在numFmt文件中未定义时返回{@code null}
     */
    public static NumFmt get(String code) {
        NumFmt fmt = null;
        for (NumFmt nf : buildInNumFmts) {
            if (nf != null && code.equals(nf.code)) {
                fmt = nf;
                break;
            }
        }
        return fmt;
    }

    /**
     * 按ID查找内置格式化对象
     *
     * @param id 内置对象ID
     * @return 对应 {@link NumFmt}，如果在numFmt文件中未定义时返回{@code null}
     */
    public static NumFmt get(int id) {
        return id >= 0 && id < buildInNumFmts.length ? buildInNumFmts[id] : null;
    }
}

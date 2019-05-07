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

import cn.ttzero.excel.util.StringUtil;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;
import java.util.Locale;

/**
 * load numFmt as default formats
 * Created by guanquan.wang at 2018-02-12 10:11
 */
public final class DefaultNumFmt {
    private static NumFmt[][] data;

    static {
        InputStream is = DefaultNumFmt.class.getClassLoader().getResourceAsStream("numFmt");
        if (is != null) {
            List<NumFmt> list = new ArrayList<>();
            int maxLen = 0;
            try (BufferedReader br = new BufferedReader(new InputStreamReader(is))) {
                String v;
                Locale locale = Locale.ROOT;
                boolean unicode = false, damaged = false;
                while ((v = br.readLine()) != null) {
                    if (StringUtil.isEmpty(v)) continue;
                    v = v.trim();
                    if (v.charAt(0) == '[') {
                        int end = v.indexOf(']');
                        if (end == -1 || end == 1) {
                            damaged = true;
                            continue;
                        }
                        String[] loc = v.substring(1, end).split("-");
                        if (loc.length < 2) { // The file is damaged
                            damaged = true;
                            continue;
                        }
                        damaged = false;
                        locale = new Locale(loc[0], loc[1]);
                        unicode = loc.length >= 3 && "unicode".equals(loc[2]);
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
                        fmt.locale = locale;
                        fmt.unicode = unicode;
                        list.add(fmt);

                        if (fmt.code.length() > maxLen) {
                            maxLen = fmt.code.length();
                        }
                    }
                }
            } catch (IOException e) {
                ; // Empty
            }

            data = new NumFmt[maxLen + 1][]; // accept zero size
            for (int i = 1; i <= maxLen; i++) {
                final int length = i;
                data[i] = list.stream()
                    .filter(o -> o.code.length() == length)
                    .sorted(Comparator.comparingInt(NumFmt::getId))
                    .toArray(NumFmt[]::new);

                if (data[i].length == 0) { // Undo empty array
                    data[i] = null;
                }
            }
        } else {
            data = new NumFmt[0][];
        }
    }

    public static int indexOf(String code) {
        NumFmt v = get(code);
        return v != null ? v.id : -1;
    }

    public static NumFmt get(String code) {
        int index = code.length();
        if (index >= data.length) return null;
        NumFmt[] array = data[index];
        if (array == null) return null;
        NumFmt v = null;
        for (NumFmt nf : array) {
            if (nf.code.equals(code)) {
                v = nf;
                break;
            }
        }
        return v;
    }

    public static class NumFmt {
        private String code;
        private int id;
        private Locale locale;
        private boolean unicode;

        public String getCode() {
            return code;
        }

        public int getId() {
            return id;
        }

        public Locale getLocale() {
            return locale;
        }

        public boolean isUnicode() {
            return unicode;
        }

        @Override
        public String toString() {
            StringBuilder buf = new StringBuilder();
            if (!Locale.ROOT.equals(locale)) {
                buf.append('[').append(locale.getLanguage()).append('-').append(locale.getCountry());
                if (unicode) {
                    buf.append('-').append("unicode");
                }
                buf.append("] ");
            }
            buf.append(id).append('=').append(code);
            return buf.toString();
        }
    }
}

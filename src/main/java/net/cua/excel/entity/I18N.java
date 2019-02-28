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

package net.cua.excel.entity;

import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.Reader;
import java.util.Arrays;
import java.util.Locale;
import java.util.Properties;

/**
 * 国际化
 * Create by guanquan.wang at 2018-10-13
 */
public class I18N {
    private Properties pro;
    private final String fn = "message._.properties";
    public I18N() {
        Locale locale = Locale.getDefault();
        pro = new Properties();
        try {
            InputStream is = I18N.class.getClassLoader().getResourceAsStream("I18N/" + fn.replace("_", locale.toLanguageTag()));
            if (is == null) {
                is = I18N.class.getClassLoader().getResourceAsStream("I18N/" + fn.replace("_", "zh-CN"));
            }
            if (is != null) {
                try (Reader reader = new InputStreamReader(is, "UTF-8")) {
                    pro.load(reader);
                }
            }
        } catch (IOException e) {
            // nothing...
        }
    }

    /**
     * get message by code
     * @param code message code
     * @return I18N string
     */
    public String get(String code) {
        if (pro != null) {
            return pro.getProperty(code);
        }
        return code;
    }

    /**
     * get message by code
     * @param code message code
     * @param other default value
     * @return I18N string
     */
    public String getOrElse(String code, String other) {
        if (pro != null) {
            return pro.getProperty(code);
        }
        return other;
    }

    /**
     * get message by code
     * @param code code
     * @param args args
     * @return I18N string
     */
    public String get(String code, String ... args) {
        if (pro == null) return code;
        String msg = pro.getProperty(code);
        char[] oldValue = msg.toCharArray();
        int[] indexs = search(oldValue);
        int len = indexs.length >= args.length ? args.length : indexs.length, size = 0;
        for (int i = 0; i < len; size += args[i++].length());
        StringBuilder buf = new StringBuilder(oldValue.length + size - (len << 1));
        buf.append(oldValue, 0, indexs[0]).append(args[0]);
        for (int i = 1; i < len; i++) {
            buf.append(oldValue, size = indexs[i - 1] + 2, indexs[i] - size).append(args[i]);
        }
        if (indexs[len - 1] + 2 < oldValue.length) {
            buf.append(oldValue, size = indexs[len - 1] + 2, oldValue.length - size);
        }
        return buf.toString();
    }

    private static int[] search(char[] value) {
        int[] indexs = new int[16];
        int n = 0;
        for (int i = 0; i < value.length - 1; i++) {
            if (value[i] == '{' && value[i + 1] == '}') {
                indexs[n++] = i++;
                if (n == indexs.length) {
                    int[] _indexs = new int[indexs.length << 1];
                    System.arraycopy(indexs, 0, _indexs, 0, indexs.length);
                    indexs = _indexs;
                }
            }
        }
        return Arrays.copyOfRange(indexs, 0, n);
    }

}

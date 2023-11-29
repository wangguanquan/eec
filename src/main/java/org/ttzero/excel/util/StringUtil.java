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

package org.ttzero.excel.util;

/**
 * 字符串工具类，提供一些简单的静态方法
 *
 * @author guanquan.wang on 2017/9/30.
 */
public class StringUtil {
    private StringUtil() { }

    /**
     * 空字符串
     */
    public final static String EMPTY = "";

    /**
     * 检查字符串是否为空字符串，当字符串{@code s}为{@code null}或{@code String.isEmpty}则返回{@code true}
     *
     * @param s 待检查字符串
     * @return {@code true}当字符串为{@code null}或{@code String.isEmpty}
     */
    public static boolean isEmpty(String s) {
        return s == null || s.isEmpty();
    }

    /**
     * 检查字符串不为空字符串，当字符串{@code s}不为{@code null}且长度大小{@code 0}则返回{@code true}
     *
     * @param s 待检查字符串
     * @return {@code true}当字符串不为{@code null}且长度大小{@code 0}
     */
    public static boolean isNotEmpty(String s) {
        return s != null && s.length() > 0;
    }

    /**
     * 查找字符串在数组中第一次出现的位置，查找是从数组头向尾逐一比较，时间复杂度{@code n}（n为数组长度），
     * 建议只应用于小数组查找，待查找字符串{@code v}可以为{@code null}，但数组不能为{@code null}
     *
     * @param array 查找源，不为{@code null}
     * @param v 待查找字符串
     * @return 如果存在则返回字符串在数组中第一次出现的下标否则返回 {@code -1}
     */
    public static int indexOf(String[] array, String v) {
        if (v != null) {
            for (int i = 0; i < array.length; i++) {
                if (v.equals(array[i])) {
                    return i;
                }
            }
        } else {
            for (int i = 0; i < array.length; i++) {
                if (array[i] == null) {
                    return i;
                }
            }
        }
        return -1;
    }

    /**
     * 首字母大写，转化是<b>强制</b>的它并不会检查空串以及第二个字符是否为大字，外部最好不要使用
     *
     * <p>注意：本方法只适用于范围为{@code [97, 122]}的{@code ASCII}值</p>
     *
     * @param key 待处理字符串
     * @return 转为首字母大写后的字符串
     */
    public static String uppFirstKey(String key) {
        char first = key.charAt(0);
        if (first >= 97 && first <= 122) {
            char[] _v = key.toCharArray();
            _v[0] -= 32;
            return new String(_v);
        }
        return key;
    }

    /**
     * 首字母小写，转化是<b>强制</b>的它并不会检查空串，外部最好不要使用
     *
     * @param key 待处理字符串
     * @return 转为首字母大写后的字符串
     */
    public static String lowFirstKey(String key) {
        char first = key.charAt(0);
        if (first >= 65 && first <= 90) {
            char[] _v = key.toCharArray();
            _v[0] += 32;
            return new String(_v);
        }
        return key;
    }

    /**
     * 将字符串转为驼峰风格，仅支持将下划线{@code '_'}风格转驼峰风格，内部不检查参数是否为{@code null}请谨慎使用
     * <blockquote><pre>
     * 转换前       | 转换后
     * ------------|------------
     * GOODS_NAME  | goodsName
     * NAME        | name
     * goods__name | goodsName
     * _goods__name| _goodsName
     * </pre></blockquote>
     * @param name 待转换字符串
     * @return 驼峰风格字符串
     */
    public static String toCamelCase(String name) {
        if (name.indexOf('_') < 0) return name.toLowerCase();
        char[] oldValues = name.toLowerCase().toCharArray();
        final int len = oldValues.length;
        int i = 1, idx = i;
        for (int n = len - 1; i < n; i++) {
            char c = oldValues[i], cc = oldValues[i + 1];
            if (c == '_') {
                if (cc == '_') continue;
                i++;
                oldValues[idx++] = cc >= 'a' && cc <= 'z' ? (char) (cc - 32) : cc;
            }
            else {
                oldValues[idx++] = c;
            }
        }
        if (i < len) oldValues[idx++] = oldValues[i];
        return new String(oldValues, 0, idx);
    }

    /**
     * 交换数组中的值，交换是<b>强制</b>的它并不会检查下标的范围，外部最好不要使用
     *
     * @param values 数组
     * @param a 指定交换下标
     * @param b 指定交换下标
     */
    public static void swap(String[] values, int a, int b) {
        String t = values[a];
        values[a] = values[b];
        values[b] = t;
    }

    /**
     * 检查字符串是否为{@code null}或空白字符
     *
     * @param cs 待检查的字符串
     * @return {@code true} 字符串为{@code null}或空白字符
     */
    public static boolean isBlank(final CharSequence cs) {
        int strLen;
        if (cs == null || (strLen = cs.length()) == 0) {
            return true;
        }
        for (int i = 0; i < strLen; i++) {
            if (!Character.isWhitespace(cs.charAt(i))) {
                return false;
            }
        }
        return true;
    }

    /**
     * 检查字符串不为{@code null}或非空白字符
     *
     * @param cs  待检查的字符串
     * @return {@code true} 字符串不为{@code null}或非空白字符
     */
    public static boolean isNotBlank(final CharSequence cs) {
        return !isBlank(cs);
    }

    /**
     * 格式化字节大小，将字节大小转为{@code kb,mb,gb}等格式
     *
     * @param size 字节大小
     * @return 格式化字符串
     */
    public static String formatBinarySize(long size) {
        long kb = 1 << 10, mb = kb << 10, gb = mb << 10;
        String s;
        if (size >= gb) s = String.format("%.2fGB", (double) size / gb);
        else if (size >= mb) s = String.format("%.2fMB", (double) size / mb);
        else if (size >= kb) s = String.format("%.2fKB", (double) size / kb);
        else s = String.format("%dB", size);
        return s.replace(".00", "");
    }


    /**
     * 毫秒时间转字符串，通常用于格式化某段代码的耗时, 如：1h:3s or 4m:1s
     *
     * @param t 毫秒时间
     * @return 格式化文本
     */
    public static String timeToString(long t) {
        int n = (int) t / 1000;
        int h = n / 3600, m = (n - h * 3600) / 60, s = n - h * 3600 - m * 60, ms = (int) (t - n * 1000);
        return (h > 0 ? h + "h" : "")
            + (m > 0 ? (h > 0 ? ":" : "") + m + "m" : "")
            + (s > 0 ? (h + m > 0 ? ":" : "") + s + "s" : "")
            + (ms > 0 ? ((h + m + s > 0 ? ":" : "") + ms + "ms") : (h + m + s > 0 ? "" : "0ms"));
    }

    /**
     * 查找某个字符{@code ch}在字符串{@code str}的位置，与{@link String#indexOf(int, int)}不同之处在于
     * 后者从开始位置查找到字符串结尾，而前者需要指定一个结束位置查找范围在{@code fromIndex}到{@code toIndex}之间
     *
     * @param str 字符串源
     * @param   ch          待查找的字符
     * @param   fromIndex   起始位置（包含）
     * @param   toIndex   结束位置（不包含）
     * @return  字符 {@code ch} 在字符串的位置，未找到时返回{@code -1}
     */
    public static int indexOf(String str, int ch, int fromIndex, int toIndex) {
        final int max = Math.min(str.length(), toIndex);
        if (fromIndex < 0) {
            fromIndex = 0;
        } else if (fromIndex >= max) {
            // Note: fromIndex might be near -1>>>1.
            return -1;
        }

        final char[] value = str.toCharArray();
        if (ch < Character.MIN_SUPPLEMENTARY_CODE_POINT) {
            // handle most cases here (ch is a BMP code point or a
            // negative value (invalid code point))
            for (int i = fromIndex; i < max; i++) {
                if (value[i] == ch) {
                    return i;
                }
            }
            return -1;
        } else {
            return indexOfSupplementary(value, ch, fromIndex, max);
        }
    }

    /**
     * UTF-8编码一个字符理论上最多占用3个字节，所以需要逐个比较每个字节，但目前为止UTF-8只使用了最多2个字节来表示世界上所有的文字，
     * 所以这里比较最多2个字节
     */
    private static int indexOfSupplementary(char[] value, int ch, int fromIndex, int toIndex) {
        if (Character.isValidCodePoint(ch)) {
            final char hi = Character.highSurrogate(ch);
            final char lo = Character.lowSurrogate(ch);
            final int max = toIndex - 1;
            for (int i = fromIndex; i < max; i++) {
                if (value[i] == hi && value[i + 1] == lo) {
                    return i;
                }
            }
        }
        return -1;
    }
}

/*
 * Copyright (c) 2019-2021, guanquan.wang@yandex.com All Rights Reserved.
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
 * string util
 * <p>
 * Created by guanquan.wang on 2017/9/30.
 */
public class StringUtil {
    private StringUtil() { }

    /**
     * Const `""` string
     */
    public final static String EMPTY = "";

    /**
     * Returns {@code true} if it is null or {@link String#length()} is {@code 0}.
     *
     * @param s string value to check
     * @return {@code true} if null or {@link String#length()} is {@code 0}, otherwise
     * {@code false}
     */
    public static boolean isEmpty(String s) {
        return s == null || s.isEmpty();
    }

    /**
     * Returns {@code true} if, and only if, {@link String#length()} greater than {@code 0}.
     *
     * @param s string value to check
     * @return {@code true} if {@link String#length()} greater than {@code 0}, otherwise
     * {@code false}
     */
    public static boolean isNotEmpty(String s) {
        return s != null && s.length() > 0;
    }

    /**
     * Returns the index within this string array of the first occurrence of
     * the specified string.
     * <p>
     * If no such string occurs in this array, then {@code -1} is returned.
     *
     * @param array array to search
     * @param v the specify string
     * @return the index of the first occurrence of the string in array, or
     *      {@code -1} if the string does not occur.
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
     * Turn the first character into an upper case. It means if the first
     * character is between 97 and 122, it will be minus {@code 32}.
     *
     * @param key a string to processor
     * @return a string witch the first character is upper case
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
     * Turn the first character into an lower case. It means if the first
     * character is between 65 and 90, it will be plus {@code 32}.
     *
     * @param key a string to processor
     * @return a string witch the first character is lower case
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
     * Convert to camel case string.
     *
     * @param name a string to processor
     * @return a camel case string
     */
    public static String toPascalCase(String name) {
        if (name.indexOf('_') < 0) return name;
        char[] oldValues = name.toLowerCase().toCharArray();
        final int len = oldValues.length;
        int i = 1, idx = i;
        for (int n = len - 1; i < n; i++) {
            char c = oldValues[i], cc = oldValues[i + 1];
            if (c == '_' && cc >= 'a' && cc <= 'z') {
                oldValues[idx++] = (char) (cc - 32);
                i++;
            }
            else {
                oldValues[idx++] = c;
            }
        }
        if (i < len) oldValues[idx++] = oldValues[i];
        return new String(oldValues, 0, idx);
    }

//    public static String append(String src, String a, int n, int origin) {
//        StringBuilder buf = new StringBuilder();
//        // Insert header
//        if (origin == -1) {
//            for (; n-- > 0; ) {
//                buf.append(a);
//            }
//            buf.append(src);
//        } else {
//            buf.append(src);
//            for (; n-- > 0; ) {
//                buf.append(a);
//            }
//        }
//
//        return buf.toString();
//    }

    /**
     * Wrap value in string array
     *
     * @param values the array to warp
     * @param a from index
     * @param b to index
     */
    public static void swap(String[] values, int a, int b) {
        String t = values[a];
        values[a] = values[b];
        values[b] = t;
    }
}

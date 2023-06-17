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
 * string util
 *
 * @author guanquan.wang on 2017/9/30.
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

    /**
     * Checks if a CharSequence is empty (""), null or whitespace only.
     *
     * @param cs  the CharSequence to check, may be null
     * @return {@code true} if the CharSequence is null, empty or whitespace only
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
     * Checks if a CharSequence is not empty (""), not null and not whitespace only.
     *
     * @param cs  the CharSequence to check, may be null
     * @return {@code true} if the CharSequence is
     *  not empty and not null and not whitespace only
     */
    public static boolean isNotBlank(final CharSequence cs) {
        return !isBlank(cs);
    }

    /**
     * long size to string
     *
     * @param size file size in bytes
     * @return String size
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
     * Time millis to String, like 1h:3s or 4m:1s
     *
     * @param t time millis
     * @return string
     */
    public static String timeToString(long t) {
        int n = (int) t / 1000;
        int h = n / 3600, m = (n - h * 3600) / 60, s = n - h * 3600 - m * 60;
        return "" + (h > 0 ? h + "h" : "")
            + (m > 0 ? (h > 0 ? ":" : "") + m + "m" : "")
            + ((h + m > 0 ? ":" : "") + s + "s");
    }

    /**
     * Returns the index within this string of the first occurrence of the
     * specified character, starting the search at the specified range.
     * <p>
     * If a character with value {@code ch} occurs in the
     * character sequence represented by this {@code String}
     * object at an index no smaller than {@code fromIndex}, then
     * the index of the first such occurrence is returned. For values
     * of {@code ch} in the range from 0 to 0xFFFF (inclusive),
     * this is the smallest value <i>k</i> such that:
     * <blockquote><pre>
     * (this.charAt(<i>k</i>) == ch) {@code &&} (<i>k</i> &gt;= fromIndex)
     * </pre></blockquote>
     * is true. For other values of {@code ch}, it is the
     * smallest value <i>k</i> such that:
     * <blockquote><pre>
     * (this.codePointAt(<i>k</i>) == ch) {@code &&} (<i>k</i> &gt;= fromIndex)
     * </pre></blockquote>
     * is true. In either case, if no such character occurs in this
     * string at or after position {@code fromIndex}, then
     * {@code -1} is returned.
     *
     * <p>
     * There is no restriction on the value of {@code fromIndex}. If it
     * is negative, it has the same effect as if it were zero: this entire
     * string may be searched. If it is greater than the length of this
     * string, it has the same effect as if it were equal to the length of
     * this string: {@code -1} is returned.
     *
     * <p>All indices are specified in {@code char} values
     * (Unicode code units).
     *
     * @param   ch          a character (Unicode code point).
     * @param   fromIndex   the index to start the search from.
     * @param   toIndex   the high endpoint (exclusive) of the search end.
     * @return  the index of the first occurrence of the character in the
     *          character sequence represented by this object that is greater
     *          than or equal to {@code fromIndex}, or {@code -1}
     *          if the character does not occur.
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
     * Handles (rare) calls of indexOf with a supplementary character.
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

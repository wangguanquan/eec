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

import java.io.BufferedWriter;
import java.io.IOException;
import java.io.Writer;

import static java.lang.Character.isHighSurrogate;
import static java.lang.Character.isLowSurrogate;
import static java.lang.Character.isSurrogate;

/**
 * Single-threaded operation stream, internal multiplexing buffer
 *
 * @author guanquan.wang on 2017/10/11.
 */
public class ExtBufferedWriter extends BufferedWriter {
    private static final int defaultCharBufferSize = 8192;

    public ExtBufferedWriter(Writer out) {
        this(out, defaultCharBufferSize);
    }

    public ExtBufferedWriter(Writer out, int sz) {
        super(out, sz);

        for (int i = 0; i < CACHE_CHAR_ARRAY.length; i++) {
            CACHE_CHAR_ARRAY[i] = new char[i + 1];
        }
    }

    private final static char[][] CACHE_CHAR_ARRAY = new char[25][];
    static final char[] MIN_INTEGER_CHARS = {'-', '2', '1', '4', '7', '4', '8', '3', '6', '4', '8'};
    static final char[] MIN_LONG_CHARS = "-9223372036854775808".toCharArray();
    private static final char[][] ESCAPE_CHARS = new char[63][];
    /**
     * Replace malformed characters
     */
    public static char MALFORMED_CHAR = 0xFFFD;

    static {
        // Fix#72 delete space escape
//        ESCAPE_CHARS[' '] = "&nbsp;".toCharArray();
        ESCAPE_CHARS['<'] = "&lt;".toCharArray();
        ESCAPE_CHARS['>'] = "&gt;".toCharArray();
        ESCAPE_CHARS['&'] = "&amp;".toCharArray();
        ESCAPE_CHARS['"'] = "&quot;".toCharArray();
    }

    /**
     * Write integer value
     *
     * @param n the integer value
     * @throws IOException if I/O error occur
     */
    public void writeInt(int n) throws IOException {
        char[] temp = toChars(n);
        write(temp);
    }

    /**
     * Write long value
     *
     * @param l the long value
     * @throws IOException if I/O error occur
     */
    public void write(long l) throws IOException {
        char[] temp = toChars(l);
        write(temp);
    }

    /**
     * Write single-precision floating-point value
     *
     * @param f the single-precision floating-point value
     * @throws IOException if I/O error occur
     */
    public void write(float f) throws IOException {
        write(Float.toString(f));
    }

    /**
     * Write as escape character
     *
     * @param c a character value
     * @throws IOException if I/O error occur
     */
    public void escapeWrite(char c) throws IOException {
        if (c > 62) {
            write(isSurrogate(c) ? MALFORMED_CHAR : c);
        }
        // Display char
        else if (c >= 32) {
            char[] ec = ESCAPE_CHARS[c];
            if (ec != null) write(ec);
            else write(c);
        } else {
            write(c == 9 || c == 10 || c == 13 ? c : MALFORMED_CHAR);
        }
    }

    /**
     * Write as escape text
     *
     * @param text string
     * @throws IOException if I/O error occur
     */
    public void escapeWrite(String text) throws IOException {
        char[] block = text.toCharArray(), ec;
        int i, last = 0, size = text.length();

        for (i = 0; i < size; i++) {
            char c = block[i];
            if (c > 62) continue;
            // Cannot display characters
            if (c < 32) {
                if (i > last) writeUTF8(block, last, i - last);
                write(c == 9 || c == 10 || c == 13 ? c : MALFORMED_CHAR);
                last = i + 1;
            }
            // html escape char
            else if ((ec = ESCAPE_CHARS[c]) != null) {
                if (i > last) writeUTF8(block, last, i - last);
                write(ec);
                last = i + 1;
            }
        }

        if (last < size) writeUTF8(block, last, i - last);
    }

    /**
     * Write double-precision floating-point value
     *
     * @param d the double-precision floating-point value
     * @throws IOException if I/O error occur
     */
    public void write(double d) throws IOException {
        write(Double.toString(d));
    }

    /**
     * Write utf-8 string
     *
     * @param  cb    A character array
     * @param  off   Offset from which to start reading characters
     * @param  len   Number of characters to write
     * @throws IOException if I/O error occur
     */
    public void writeUTF8(char[] cb, int off, int len) throws IOException {
        if (len <= 0) return;
        int end = off + len, i = lookupMalformedUTF8Char(cb, off, end);
        if (i >= 0) {
            cb[i++] = MALFORMED_CHAR;
            for (; (i = lookupMalformedUTF8Char(cb, i, end)) >= 0; cb[i++] = MALFORMED_CHAR);
        }
        super.write(cb, off, len);
    }

    public static char[] toChars(int i) {
        if (i == Integer.MIN_VALUE)
            return MIN_INTEGER_CHARS;
        int size = stringSize(i);
        getChars(i, size, CACHE_CHAR_ARRAY[size - 1]);
        return CACHE_CHAR_ARRAY[size - 1];
    }


    private final static int[] sizeTable = {9, 99, 999, 9999, 99999, 999999, 9999999,
        99999999, 999999999, Integer.MAX_VALUE};

    // Requires positive x
    public static int stringSize(int x) {
        boolean negative = x < 0;
        if (negative) x = ~x + 1;
        int l;
        for (int i = 0; ; i++)
            if (x <= sizeTable[i]) {
                l = i + 1;
                break;
            }
        return negative ? l + 1 : l;
    }

    public static void getChars(int i, int maxIndex, char[] buf) {
        if (i == 0) {
            buf[maxIndex - 1] = '0';
            return;
        }
        boolean negative = i < 0;
        if (negative) i = -i;
        for (; i > 0; buf[--maxIndex] = (char) ((i % 10) + '0'), i /= 10);
        if (negative) buf[--maxIndex] = '-';
    }

    public static char[] toChars(long i) {
        if (i == Long.MIN_VALUE)
            return MIN_LONG_CHARS;
        int size = stringSize(i);
        getChars(i, size, CACHE_CHAR_ARRAY[size - 1]);
        return CACHE_CHAR_ARRAY[size - 1];
    }

    // Requires positive x
    public static int stringSize(long x) {
        boolean negative = x < 0;
        if (negative) x = ~x + 1;
        int l = 0, i = 1;
        long p = 10;
        for (; i < 19; i++) {
            if (x < p) {
                l = i;
                break;
            }
            p = 10 * p;
        }
        if (i >= 19) l = 19;
        return negative ? l + 1 : l;
    }

    public static void getChars(long i, int maxIndex, char[] buf) {
        if (i == 0) {
            buf[maxIndex - 1] = '0';
            return;
        }
        boolean negative = i < 0;
        if (negative) i = -i;
        for (; i > 0; buf[--maxIndex] = (char) ((i % 10) + '0'), i /= 10);
        if (negative) buf[--maxIndex] = '-';
    }

    // Find malformed characters and return to their location, -1 means OK
    static int lookupMalformedUTF8Char(char[] cb, int from, int to) {
        for (char c; from < to; from++) {
            c = cb[from];
            if (!isSurrogate(c)) continue;
            if (isHighSurrogate(c)) {
                if (to - from < 2 || !isLowSurrogate(cb[from + 1])) return from;
                from++;
            }
            else if (isLowSurrogate(c)) return from;
        }
        return -1;
    }
}

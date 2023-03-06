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

/**
 * Single-threaded operation stream, internal multiplexing buffer
 *
 * @author guanquan.wang on 2017/10/11.
 */
public class ExtBufferedWriter extends BufferedWriter {
    private static final int defaultCharBufferSize = 8192;
//    private long cn; // Char length

    public ExtBufferedWriter(Writer out) {
        this(out, defaultCharBufferSize);
    }

    public ExtBufferedWriter(Writer out, int sz) {
        super(out, sz);

        for (int i = 0; i < cache_char_array.length; i++) {
            cache_char_array[i] = new char[i + 1];
        }
    }

    private final char[][] cache_char_array = new char[25][];
    static final char[] MIN_INTEGER_CHARS = {'-', '2', '1', '4', '7', '4', '8', '3', '6', '4', '8'};
    static final char[] MIN_LONG_CHARS = "-9223372036854775808".toCharArray();
    private static final char[][] ESCAPE_CHARS = new char[63][];

    static {
//        for (int i = 1; i < 32; i++) {
//            ESCAPE_CHARS[i] = ("&#" + i + ";").toCharArray();
//        }
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
            write(c);
        }
        // Display char
        else if (c >= 32) {
            char[] entity = ESCAPE_CHARS[c];
            if (entity != null) write(entity);
            else write(c);
        }
    }

    /**
     * Write as escape text
     *
     * @param text string
     * @throws IOException if I/O error occur
     */
    public void escapeWrite(String text) throws IOException {
        char[] block = text.toCharArray();
        int i;
        int last = 0;
        int size = text.length();

        for (i = 0; i < size; i++) {
            char c = block[i];
            if (c > 62) continue;
            // UnDisplay char
//            if (c < 32) {
//                write(block, last, i - last);
//                last = i + 1;
//                continue;
//            }
            // html escape char
            char[] entity = ESCAPE_CHARS[c];

            if (entity != null) {
                write(block, last, i - last);
                write(entity);
                last = i + 1;
            }
        }

        if (last < size) {
            write(block, last, i - last);
        }
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
//
//    @Override
//    public void write(int c) throws IOException {
//        super.write(c);
//        cn++;
//    }
//
//    @Override
//    public void write(char[] cbuf, int off, int len) throws IOException {
//        super.write(cbuf, off, len);
//        cn += len;
//    }
//
//    @Override
//    public void write(String s, int off, int len) throws IOException {
//        super.write(s, off, len);
//        cn += len;
//    }
//
//    /**
//     * Returns the number of characters that have been written
//     *
//     * @return number of characters
//     */
//    public long getWrittenChars() {
//        return cn;
//    }

    public char[] toChars(int i) {
        if (i == Integer.MIN_VALUE)
            return MIN_INTEGER_CHARS;
        int size = stringSize(i);
        getChars(i, size, cache_char_array[size - 1]);
        return cache_char_array[size - 1];
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

    static void getChars(int i, int index, char[] buf) {
        int q, r;
        int charPos = index;
        char sign = 0;

        if (i < 0) {
            sign = '-';
            i = -i;
        }

        // Generate two digits per iteration
        while (i >= 65536) {
            q = i / 100;
            // really: r = i - (q * 100);
            r = i - ((q << 6) + (q << 5) + (q << 2));
            i = q;
            buf[--charPos] = digitOnes[r];
            buf[--charPos] = digitTens[r];
        }

        // Fall thur to fast mode for smaller numbers
        // assert(i <= 65536, i);
        for (; ; ) {
            q = (i * 52429) >>> (16 + 3);
            r = i - ((q << 3) + (q << 1));  // r = i-(q*10) ...
            buf[--charPos] = digits[r];
            i = q;
            if (i == 0) break;
        }
        if (sign != 0) {
            buf[--charPos] = sign;
        }
    }

    public char[] toChars(long i) {
        if (i == Long.MIN_VALUE)
            return MIN_LONG_CHARS;
        int size = stringSize(i);
        getChars(i, size, cache_char_array[size - 1]);
        return cache_char_array[size - 1];
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

    /**
     * Places characters representing the integer i into the
     * character array buf. The characters are placed into
     * the buffer backwards starting with the least significant
     * digit at the specified index (exclusive), and working
     * backwards from there.
     * <p>
     * Will fail if i == Long.MIN_VALUE
     */
    static void getChars(long i, int index, char[] buf) {
        long q;
        int r;
        int charPos = index;
        char sign = 0;

        if (i < 0) {
            sign = '-';
            i = -i;
        }

        // Get 2 digits/iteration using longs until quotient fits into an int
        while (i > Integer.MAX_VALUE) {
            q = i / 100;
            // really: r = i - (q * 100);
            r = (int) (i - ((q << 6) + (q << 5) + (q << 2)));
            i = q;
            buf[--charPos] = digitOnes[r];
            buf[--charPos] = digitTens[r];
        }

        // Get 2 digits/iteration using ints
        int q2;
        int i2 = (int) i;
        while (i2 >= 65536) {
            q2 = i2 / 100;
            // really: r = i2 - (q * 100);
            r = i2 - ((q2 << 6) + (q2 << 5) + (q2 << 2));
            i2 = q2;
            buf[--charPos] = digitOnes[r];
            buf[--charPos] = digitTens[r];
        }

        // Fall thur to fast mode for smaller numbers
        // assert(i2 <= 65536, i2);
        for (; ; ) {
            q2 = (i2 * 52429) >>> (16 + 3);
            r = i2 - ((q2 << 3) + (q2 << 1));  // r = i2-(q2*10) ...
            buf[--charPos] = digits[r];
            i2 = q2;
            if (i2 == 0) break;
        }
        if (sign != 0) {
            buf[--charPos] = sign;
        }
    }

    public final static char[] digitTens = {
        '0', '0', '0', '0', '0', '0', '0', '0', '0', '0',
        '1', '1', '1', '1', '1', '1', '1', '1', '1', '1',
        '2', '2', '2', '2', '2', '2', '2', '2', '2', '2',
        '3', '3', '3', '3', '3', '3', '3', '3', '3', '3',
        '4', '4', '4', '4', '4', '4', '4', '4', '4', '4',
        '5', '5', '5', '5', '5', '5', '5', '5', '5', '5',
        '6', '6', '6', '6', '6', '6', '6', '6', '6', '6',
        '7', '7', '7', '7', '7', '7', '7', '7', '7', '7',
        '8', '8', '8', '8', '8', '8', '8', '8', '8', '8',
        '9', '9', '9', '9', '9', '9', '9', '9', '9', '9',
    };

    public final static char[] digitOnes = {
        '0', '1', '2', '3', '4', '5', '6', '7', '8', '9',
        '0', '1', '2', '3', '4', '5', '6', '7', '8', '9',
        '0', '1', '2', '3', '4', '5', '6', '7', '8', '9',
        '0', '1', '2', '3', '4', '5', '6', '7', '8', '9',
        '0', '1', '2', '3', '4', '5', '6', '7', '8', '9',
        '0', '1', '2', '3', '4', '5', '6', '7', '8', '9',
        '0', '1', '2', '3', '4', '5', '6', '7', '8', '9',
        '0', '1', '2', '3', '4', '5', '6', '7', '8', '9',
        '0', '1', '2', '3', '4', '5', '6', '7', '8', '9',
        '0', '1', '2', '3', '4', '5', '6', '7', '8', '9',
    };

    public final static char[] digits = {
        '0', '1', '2', '3', '4', '5',
        '6', '7', '8', '9', 'a', 'b',
        'c', 'd', 'e', 'f', 'g', 'h',
        'i', 'j', 'k', 'l', 'm', 'n',
        'o', 'p', 'q', 'r', 's', 't',
        'u', 'v', 'w', 'x', 'y', 'z'
    };

    public final static char[] digits_uppercase = {
        '0', '1', '2', '3', '4', '5',
        '6', '7', '8', '9', 'A', 'B',
        'C', 'D', 'E', 'F', 'G', 'H',
        'I', 'J', 'K', 'L', 'M', 'N',
        'O', 'P', 'Q', 'R', 'S', 'T',
        'U', 'V', 'W', 'X', 'Y', 'Z'
    };

}

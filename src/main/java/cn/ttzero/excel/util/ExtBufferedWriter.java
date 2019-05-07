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

package cn.ttzero.excel.util;

import java.io.BufferedWriter;
import java.io.IOException;
import java.io.Writer;

/**
 * 单线种操作流，内部复用buffer
 * Created by guanquan.wang on 2017/10/11.
 */
public class ExtBufferedWriter extends BufferedWriter {
    private static int defaultCharBufferSize = 8192;

    public ExtBufferedWriter(Writer out) {
        this(out, defaultCharBufferSize);
    }

    public ExtBufferedWriter(Writer out, int sz) {
        super(out, sz);

        for (int i = 0; i < cache_char_array.length; i++) {
            cache_char_array[i] = new char[i + 1];
        }
    }

    private char[][] cache_char_array = new char[25][];
    private static final char[] MIN_INTEGER_CHARS = {'-', '2', '1', '4', '7', '4', '8', '3', '6', '4', '8'};
    private static final char[] MIN_LONG_CHARS = "-9223372036854775808".toCharArray();
    private static final char[][] ESCAPE_CHARS = new char[63][];

    static {
//        for (int i = 1; i < 32; i++) {
//            ESCAPE_CHARS[i] = ("&#" + i + ";").toCharArray();
//        }
        ESCAPE_CHARS[' '] = "&nbsp;".toCharArray();
        ESCAPE_CHARS['<'] = "&lt;".toCharArray();
        ESCAPE_CHARS['>'] = "&gt;".toCharArray();
        ESCAPE_CHARS['&'] = "&amp;".toCharArray();
        ESCAPE_CHARS['"'] = "&quot;".toCharArray();
    }

    /**
     * @param n
     * @return
     */
    public void writeInt(int n) throws IOException {
        char[] temp = toChars(n);
        write(temp);
    }

    public void write(long l) throws IOException {
        char[] temp = toChars(l);
        write(temp);
    }

    public void write(float f) throws IOException {
        write(Float.toString(f));
    }

    /**
     * escape text
     *
     * @param text string
     * @throws IOException
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
            if (c < 32) {
                write(block, last, i - last);
                last = i + 1;
                continue;
            }
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
     * @param d
     * @return
     */
    public void write(double d) throws IOException {
        write(Double.toString(d));
    }

    private char[] toChars(int i) {
        if (i == Integer.MIN_VALUE)
            return MIN_INTEGER_CHARS;
        int size = (i < 0) ? stringSize(-i) + 1 : stringSize(i);
        getChars(i, size, cache_char_array[size - 1]);
        return cache_char_array[size - 1];
    }


    private final static int[] sizeTable = {9, 99, 999, 9999, 99999, 999999, 9999999,
        99999999, 999999999, Integer.MAX_VALUE};

    // Requires positive x
    public static int stringSize(int x) {
        for (int i = 0; ; i++)
            if (x <= sizeTable[i])
                return i + 1;
    }

    private static void getChars(int i, int index, char[] buf) {
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

    private char[] toChars(long i) {
        if (i == Long.MIN_VALUE)
            return MIN_LONG_CHARS;
        int size = (i < 0) ? stringSize(-i) + 1 : stringSize(i);
        getChars(i, size, cache_char_array[size - 1]);
        return cache_char_array[size - 1];
    }

    // Requires positive x
    public static int stringSize(long x) {
        long p = 10;
        for (int i = 1; i < 19; i++) {
            if (x < p)
                return i;
            p = 10 * p;
        }
        return 19;
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
    private static void getChars(long i, int index, char[] buf) {
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

    private final static char[] digitTens = {
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

    private final static char[] digitOnes = {
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

    private final static char[] digits = {
        '0', '1', '2', '3', '4', '5',
        '6', '7', '8', '9', 'a', 'b',
        'c', 'd', 'e', 'f', 'g', 'h',
        'i', 'j', 'k', 'l', 'm', 'n',
        'o', 'p', 'q', 'r', 's', 't',
        'u', 'v', 'w', 'x', 'y', 'z'
    };

}

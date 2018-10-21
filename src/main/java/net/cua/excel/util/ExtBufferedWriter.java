package net.cua.excel.util;

import java.io.BufferedWriter;
import java.io.IOException;
import java.io.Writer;

/**
 * 单线种操作流，内部复用buffer
 * Created by guanquan.wang at 2017/10/11.
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

        long p = 10;
        for (int i = 0; i < longSizeTable.length; i++) {
            longSizeTable[i] = p;
            p = 10 * p;
        }
    }

    char[][] cache_char_array = new char[25][];
    static final char[] MIN_INTEGER_CHARS = {'-', '2', '1', '4', '7', '4', '8', '3', '6', '4', '8'};
    static final char[] MIN_LONG_CHARS = "-9223372036854775808".toCharArray();

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
     * @param d
     * @return
     */
    public void write(double d) throws IOException {
        write(Double.toString(d));
    }

    char[] toChars(int i) {
        if (i == Integer.MIN_VALUE)
            return MIN_INTEGER_CHARS;
        int size = (i < 0) ? stringSize(-i) + 1 : stringSize(i);
        getChars(i, size, cache_char_array[size - 1]);
        return cache_char_array[size - 1];
    }


    final static int[] sizeTable = {9, 99, 999, 9999, 99999, 999999, 9999999,
            99999999, 999999999, Integer.MAX_VALUE};

    // Requires positive x
    public static int stringSize(int x) {
        for (int i = 0; ; i++)
            if (x <= sizeTable[i])
                return i + 1;
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

    char[] toChars(long i) {
        if (i == Long.MIN_VALUE)
            return MIN_LONG_CHARS;
        int size = (i < 0) ? stringSize(-i) + 1 : stringSize(i);
        getChars(i, size, cache_char_array[size - 1]);
        return cache_char_array[size - 1];
    }

    long[] longSizeTable = new long[19];

    // Requires positive x
    int stringSize(long x) {
        for (int i = 0; i < 20; i++) {
            if (x < longSizeTable[i])
                return i + 1;
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

    final static char[] digitTens = {
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

    final static char[] digitOnes = {
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

    final static char[] digits = {
            '0', '1', '2', '3', '4', '5',
            '6', '7', '8', '9', 'a', 'b',
            'c', 'd', 'e', 'f', 'g', 'h',
            'i', 'j', 'k', 'l', 'm', 'n',
            'o', 'p', 'q', 'r', 's', 't',
            'u', 'v', 'w', 'x', 'y', 'z'
    };
}

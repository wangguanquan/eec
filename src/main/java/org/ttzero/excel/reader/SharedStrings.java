/*
 * Copyright (c) 2017-2018, guanquan.wang@hotmail.com All Rights Reserved.
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

package org.ttzero.excel.reader;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.Closeable;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.Reader;
import java.nio.charset.StandardCharsets;
import java.util.Arrays;

import static java.lang.Character.highSurrogate;
import static java.lang.Character.isBmpCodePoint;
import static java.lang.Character.isValidCodePoint;
import static java.lang.Character.lowSurrogate;
import static java.lang.Integer.numberOfTrailingZeros;
import static org.ttzero.excel.manager.Const.Limit.MAX_CHARACTERS_PER_CELL;
import static org.ttzero.excel.util.ExtBufferedWriter.MALFORMED_CHAR;
import static org.ttzero.excel.util.StringUtil.EMPTY;

/**
 * Read sharedString data
 * <p>
 * This record contains a list of all strings used anywhere in the workbook.
 * Each string occurs only once. The workbook uses indexes into the list to reference the strings
 *
 * @author guanquan.wang at 2018-09-27 14:28
 */
public class SharedStrings implements Closeable {
    private final Logger LOGGER = LoggerFactory.getLogger(getClass());

    /**
     * The maximum capacity, used if a higher value is implicitly specified
     * by either of the constructors with arguments.
     * MUST be a power of two <= 1<<36.
     */
    static final int MAXIMUM_CAPACITY = 1 << 20;

    /**
     * Constructs a SharedStrings containing the elements of the
     * specified data array
     *
     * @param data the shared strings
     */
    public SharedStrings(String[] data) {
        max = data.length;
        offset_forward = 0;
        status = 1;
        if (max <= 512) {
            forward = new String[max];
            System.arraycopy(data, offset_forward, forward, 0, max);
            limit_forward = max;
        } else {
            page = (max + 1) >> 1;
            status <<= 1;
            forward = new String[page];
            limit_forward = page;
            System.arraycopy(data, offset_forward, forward, 0, limit_forward);
            offset_backward = page;
            limit_backward = max - page;
            backward = new String[limit_backward];
            System.arraycopy(data, offset_backward, backward, 0, limit_backward);
        }
    }

    /**
     * Constructs a SharedString with the xml path, please call
     * {@link SharedStrings#load()} after instance
     *
     * @param is   the xml file path
     * @param cacheSize the number of word per load
     * @param hotSize   the number of high frequency word
     */
    public SharedStrings(InputStream is, int cacheSize, int hotSize) {
        this.reader = new InputStreamReader(is, StandardCharsets.UTF_8);
        if (cacheSize > 0) {
            this.page = tableSizeFor(cacheSize);
        }
        this.hotSize = hotSize;
    }

    /**
     * Constructs a SharedStrings with a {@link IndexSharedStringTable}
     *
     * @param sst {@link IndexSharedStringTable}
     * @param cacheSize the number of word per load
     * @param hotSize   the number of high frequency word
     * @throws IOException if I/O error occur.
     */
    public SharedStrings(IndexSharedStringTable sst, int cacheSize, int hotSize) throws IOException {
        this.sst = sst;
        max = sst.size();
        if (cacheSize > 0) {
            this.page = tableSizeFor(cacheSize);
        }
        this.hotSize = hotSize;
        init();
        // Load forward
        limit_forward = sst.get(offset_forward = 0, forward);
    }

    /**
     * Storage the new load data
     */
    private String[] forward;
    /**
     * Copy data to this area when the forward area is missing
     */
    private String[] backward;
    /**
     * Number of word per load
     */
    private int page;
    /**
     * The word total
     */
    private int max = -1, offsetM = 0;
    /**
     * The forward offset
     */
    private int offset_forward = -1;
    /**
     * The backward offset
     */
    private int offset_backward = -1;
    /**
     * The forward limit
     */
    private int limit_forward;
    /**
     * The backward limit
     */
    private int limit_backward;
    /**
     * A tester of SharedString's cache
     */
    private Tester tester = null;
    /**
     * High frequency word
     */
    private Cache<Integer, String> hot;
    /**
     * Size of hot
     */
    private int hotSize;
    /**
     * Main reader
     */
    private Reader reader;
    /**
     * Buffered
     */
    private char[] cb;
    /**
     * length of cb[]
     */
    private int nChar, length;
    /**
     * Shared string table
     */
    private IndexSharedStringTable sst;
    /**
     * 0: empty
     * 1: forward only
     * 2: forward + backward
     * 4: large model/unknown size
     */
    private int status;
    /**
     * Buffer
     */
    StringBuilder buf = null;

    // For debug
    private int total, total_forward, total_backward, total_hot, total_sst;

    /**
     * @return the shared string unique count
     * -1 if unknown size
     */
    public int size() {
        return max;
    }

    /**
     * Returns a power of two size for the given target capacity.
     *
     * @param cap the custom buffer size
     * @return Returns a power of two size
     */
    public static int tableSizeFor(int cap) {
        int n = cap - 1;
        n |= n >>> 1;
        n |= n >>> 2;
        n |= n >>> 4;
        n |= n >>> 8;
        n |= n >>> 16;
        return (n < 64) ? 64 : (n >= MAXIMUM_CAPACITY) ? MAXIMUM_CAPACITY : n + 1;
    }

    /**
     * Load the sharedString.xml file and instance word cache
     *
     * @return the {@code SharedStrings}
     * @throws IOException if io error occur
     */
    public SharedStrings load() throws IOException {
        // Get unique count
        max = uniqueCount();
        LOGGER.debug("Size of SharedString: {}", max);
        //
        init();
        return this;
    }

    /* */
    private void init() throws IOException {
        status = 1;
        // Unknown size or greater than {@code 8192}
        if (max < 0 || max > 1 << 14) {
            if (page <= 0) page = 16;
            status <<= 2;
            forward = new String[page];
            backward = new String[page];

            // Cache 8KB binary, it will store 1^16 strings.
            tester = new Tester.BinaryTester(max > 0 ? Math.min(max, 1 << 16) : 1 << 16);

            if (hotSize > 0) hot = FixSizeLRUCache.create(hotSize);
            else hot = FixSizeLRUCache.create();
            // Instance the SharedStringTable
            if (sst == null) {
                sst = new IndexSharedStringTable();
                sst.setShortSectorSize(numberOfTrailingZeros(page));
            }
        }
        else if (max > 512) {
            status <<= 1;
            page = (max + 1) >> 1;
            forward = new String[page];
            backward = new String[page];
        }
        else {
            forward = new String[page = Math.max(16, max)];
        }
    }

    /**
     * Getting the unique strings count in SharedStringTable
     *
     * @return the unique strings count
     * @throws IOException if I/O error occur
     */
    private int uniqueCount() throws IOException {
        int off = -1;
        cb = new char[1 << 12];
        length = reader.read(cb);

        // Empty Shared String Table
        if (length <= 0) return status = 0;

        String line = new String(cb, 0, Math.min(256, length));
        String uniqueCount = " uniqueCount=";
        int index = line.indexOf(uniqueCount)
            , end = index > 0 ? line.indexOf('"', index += (uniqueCount.length() + 1)) : -1;
        if (end > 0) {
            off = Integer.parseInt(line.substring(index, end));
        }
        else {
            String count = " count=";
            index = line.indexOf(count);
            end = index > 0 ? line.indexOf('"', index += (count.length() + 1)) : -1;
            if (end > 0) {
                off = Integer.parseInt(line.substring(index, end));
            }
        }

        if (end > 0) nChar = end + 1;

        return off;
    }

    /**
     * Getting the strings value by index
     *
     * @param index the index of SharedStringTable
     * @return string
     */
    public String get(int index) {
//        checkBound(index);
        total++;

        // Load first
        if (offset_forward == -1) {
            offset_forward = index / page * page;

            readMore();
        }

        String value = null;
        // Find in forward
        if (forwardRange(index)) {
            value = forward[index - offset_forward];
            total_forward++;
            if (test(index)) hot.put(index, value);
            return value;
        }

        if (status == 1)
            throw new IndexOutOfBoundsException("Index: " + index + ", Size: " + max);

        // Find in backward
        if (backwardRange(index)) {
            value = backward[index - offset_backward];
            total_backward++;
            if (test(index)) hot.put(index, value);
            return value;
        }

        // Find in hot cache
        if (status == 4) {
            value = hot.get(index);
        }

        // Can't find in memory cache
        if (value == null) {
            if (status == 2 && offset_backward > -1)
                throw new IndexOutOfBoundsException("Index: " + index + ", Size: " + max);

            copyToBackward();
            // reload data
            offset_forward = index / page * page;
            forward[0] = null;
            if (status == 4 && index < sst.size()) {
                try {
                    // Load from SharedStringTable
                    limit_forward = sst.get(offset_forward, forward);
                } catch (IOException e) {
                    throw new ExcelReadException(e);
                }
                total_sst++;
            } else {
                readMore();
                total_forward++;
            }
            if (forward[0] == null) {
                throw new IndexOutOfBoundsException("Index: " + index + ", Size: " + max);
            }
            value = forward[index - offset_forward];
            if (test(index)) hot.put(index, value);
        } else {
            total_hot++;
        }

        return value;
    }

    // Check the forward range
    private boolean forwardRange(int index) {
        return offset_forward >= 0 && offset_forward <= index
            && offset_forward + limit_forward > index;
    }

    // Check the backward range
    private boolean backwardRange(int index) {
        return offset_backward >= 0 && offset_backward <= index
            && offset_backward + limit_backward > index;
    }

    // Check the current index has been loaded twice
    private boolean test(int index) {
        return status == 4 && tester.test(index);
    }

    private void copyToBackward() {
        String[] tmp = backward;
        backward = forward;
        forward = tmp;
        offset_backward = offset_forward;
        limit_backward = limit_forward;
    }

    /**
     * Load string record from xml
     */
    protected void readMore() {
        int index = offset_forward / page;
        try {
            // Read xml file string value into IndexSharedStringTable
            for (int n = index - offsetM; n-- >= 0; ) {
                if (offset_backward == -1 && limit_forward > 0) {
                    copyToBackward();
                    offset_backward = 0;
                }
                readData();
            }
        } catch (IOException e) {
            throw new ExcelReadException(e);
        }
    }

    /**
     * Read data from main reader
     * forward only
     *
     * @return the word count
     * @throws IOException if I/O error occur
     */
    protected int readData() throws IOException {
        // Read forward area data
        int n = 0, len = length, offset;
        for (; ;) {
            int len0 = len - 3, len1 = len0 - 1;
            int[] t = findT(cb, nChar, len, len0, len1, n);

            nChar = t[0];
            limit_forward = n = t[1];

            // A page Or EOF
            if (n == page || len < cb.length && nChar == len - 6) {
                ++offsetM; // out of index range
                break;
            }

            // If cell value(character value) length greater than buffer size
            if (nChar == 0) {
                cb = Arrays.copyOf(cb, cb.length << 1);
                offset = len;
            }
            else if (nChar < len) System.arraycopy(cb, nChar, cb, 0, offset = len - nChar);
            else offset = 0;

            // Read more
            if ((len = reader.read(cb, offset, cb.length - offset)) <= 0) break;
            length = len += offset;
            nChar = 0;
        }

        // Reset totals when unknown size
        if (max < n) {
            max = offsetM * page + n;
        }
        return n; // Returns the word count
    }

    // [0]: nChar
    // [1]: number of string
    private int[] findT(char[] cb, int nChar, int length, int len0, int len1, int n) throws IOException {
        int cursor;
        for (; nChar < length && n < page; ) {
            cursor = nChar;
            // find the tag `<si>` or tag `<si/>`
            for (; nChar < len0 && cb[nChar] != '<'; ++nChar) ;
            // Empty
            if (nChar < len0 && cb[nChar + 1] == 's' && cb[nChar + 2] == 'i' && (cb[nChar + 3] == '>' || cb[nChar + 3] == '/' && cb[nChar + 4] == '>')) {
                if (cb[nChar + 3] == '/') {
                    forward[n++] = EMPTY;
                    if (status == 4) sst.push(forward[n - 1]);
                    nChar += 5;
                    continue;
                } else nChar += 4;
            }

            int[] subT = subT(cb, nChar, len0, len1);
            int a = subT[0];
            if (a == -1) break;
            nChar = subT[1];

            String tmp = escape(cb, a, nChar);

             // Skip the end tag of 't'
            nChar += 4;

            // Test the next tag
            if (cb[nChar] != '<') for (; nChar < len1 && (cb[nChar] != '<'); ++nChar);

            // End of <si>
            if (nChar < len1 && cb[nChar + 1] == '/' && cb[nChar + 2] == 's' && cb[nChar + 3] == 'i' && cb[nChar + 4] == '>') {
                forward[n++] = tmp;
                if (status == 4) sst.push(forward[n - 1]);
                nChar += 5;
            } else {
                int t = nChar;
                // Find the end tag of 'si'
                for (; nChar < len1 && (cb[nChar] != '<' || cb[nChar + 1] != '/'
                    || cb[nChar + 2] != 's' || cb[nChar + 3] != 'i' || cb[nChar + 4] != '>'); ++nChar);
                if (nChar >= len1) {
                    nChar = cursor;
                    break;
                }
                int end = nChar;
                nChar = t;
                boolean shouldClear = true;
                // Loop and join
                for (; ; ) {
                    subT = subT(cb, nChar, end, end - 1);
                    a = subT[0];
                    if (a == -1) break;
                    nChar = subT[1];

                    if (buf == null) {
                        shouldClear = false;
                        buf = new StringBuilder(tmp);
                    }
                    else if (shouldClear) {
                        shouldClear = false;
                        buf.delete(0, buf.length());
                        buf.append(tmp);
                    }
                    buf.append(escape(cb, a, nChar));
                    nChar += 4;
                }
                forward[n++] = shouldClear ? tmp : buf.toString();
                if (status == 4) sst.push(forward[n - 1]);
                nChar = end + 5;
            }

            // An integral page records
            if (n == page) break;
        }
        // DEBUG the last character
//        LOGGER.info("---------{}---------", new String(cb, nChar, length - nChar));
        return new int[] { nChar, n };
    }

    // Returns the index round of <t>
    private int[] subT(char[] cb, int nChar, int len0, int len1) {
        do {
            // The next tag
            for (; nChar < len0 && cb[nChar] != '<'; ++nChar) ;

            if (nChar >= len1) return new int[] { -1 };

            // Ignore <rPh> translate
            if (cb[nChar + 1] == 'r' && cb[nChar + 2] == 'P' && cb[nChar + 3] == 'h' && (cb[nChar + 4] == '>' || cb[nChar + 4] == ' ')) {
                int a = nChar + 5;
                for (int len = len1 - 2; a < len && cb[a] != '<' || cb[a + 1] != '/' || cb[a + 2] != 'r'
                        || cb[a + 3] != 'P' || cb[a + 4] != 'h' || cb[a + 5] != '>'; ++a)
                    ;
                if (a >= len1 - 2) return new int[] { -1 };
                nChar = a + 6;
            } else break;
        } while (nChar < len1);

        // Empty si
        if (nChar < len1 && cb[nChar + 1] == '/' && cb[nChar + 2] == 's' && cb[nChar + 3] == 'i' && cb[nChar + 4] == '>') {
            // It will skip the </t> tag, so here you need to go back 4 characters in reverse
            return new int[] { nChar - 4, nChar - 4 };
        }

        for (; nChar < len0 && (cb[nChar] != '<' || cb[nChar + 1] != 't'
            || cb[nChar + 2] != '>' && cb[nChar + 2] != ' ' && cb[nChar + 2] != '/'); ++nChar)
            ;
        if (nChar >= len0) return new int[] { -1 }; // Not found
        // Empty tag
        if (cb[nChar + 2] == '/' && cb[nChar + 3] == '>') return new int[] { nChar, nChar };
        int a = nChar += 3;
        if (cb[nChar - 1] == ' ') { // space="preserve"
            for (; nChar < len0 && cb[nChar++] != '>'; ) ;
            if (nChar >= len0) return new int[] { -1 }; // Not found
            a = nChar;
        }
        for (; nChar < len1 && (cb[nChar] != '<' || cb[nChar + 1] != '/'
            || cb[nChar + 2] != 't' || cb[nChar + 3] != '>'); ++nChar)
            ;
        if (nChar >= len1) return new int[] { -1 }; // Not found

        return new int[] { a, nChar };
    }

    // Buffer cache (Maximum 64K)
    private static char[] charBuffer = {};

    /**
     * escape
     *
     * @param cb source char buffer
     * @param from starting position in the source array.
     * @param to ending position in the source array.
     * @return Escape xml string
     */
    public static String escape(char[] cb, int from, int to) {
        int n = to - from;
        if (n == 0) return EMPTY;
        int idx_38 = indexOf(cb, '&', from, to)
            , idx_59 = idx_38 > -1 && idx_38 < to ? indexOf(cb, ';', idx_38 + 1, Math.min(idx_38 + 9, to)) : -1;

        if (idx_38 < from || idx_38 >= idx_59 || idx_59 > to) return new String(cb, from, to - from);

        char[] buf;
        if (n <= charBuffer.length) buf = charBuffer;
        else if (n <= MAX_CHARACTERS_PER_CELL) charBuffer = buf = new char[Math.min(n + 100, MAX_CHARACTERS_PER_CELL)];
        else buf = new char[n];

        int offset = 0;
        do {
            System.arraycopy(cb, from, buf, offset, n = idx_38 - from);
            offset += n;
            // ASCII
            if (cb[idx_38 + 1] == '#') {
                char c = cb[idx_38 + 2];
                if (c == 'x') offset += toChars(toIntH(cb, idx_38 + 3, idx_59), buf, offset);
                else if (c >= '0' && c <= '9') offset += toChars(toInt(cb, idx_38 + 2, idx_59), buf, offset);
                else {
                    System.arraycopy(cb, idx_38, buf, offset, n = idx_59 - idx_38 + 1);
                    offset += n;
                }
            }
            // desc
            else {
                n = idx_59 - idx_38 - 1;
                if (n == 2 && cb[idx_38 + 1] == 'l' && cb[idx_38 + 2] == 't') buf[offset++] = '<';
                else if (n == 2 && cb[idx_38 + 1] == 'g' && cb[idx_38 + 2] == 't') buf[offset++] = '>';
                else if (n == 3 && cb[idx_38 + 1] == 'a' && cb[idx_38 + 2] == 'm' && cb[idx_38 + 3] == 'p') buf[offset++] = '&';
                else if (n == 4 && cb[idx_38 + 1] == 'n' && cb[idx_38 + 2] == 'b' && cb[idx_38 + 3] == 's' && cb[idx_38 + 4] == 'p') buf[offset++] = ' ';
                else if (n == 4 && cb[idx_38 + 1] == 'q' && cb[idx_38 + 2] == 'u' && cb[idx_38 + 3] == 'o' && cb[idx_38 + 4] == 't') buf[offset++] = '"';
                else if (n == 4 && cb[idx_38 + 1] == 'a' && cb[idx_38 + 2] == 'p' && cb[idx_38 + 3] == 'o' && cb[idx_38 + 4] == 's') buf[offset++] = '\'';
                else {
                    System.arraycopy(cb, idx_38, buf, offset, n = idx_59 - idx_38 + 1);
                    offset += n;
                }
            }
            from = ++idx_59;
            idx_59 = (idx_38 = indexOf(cb, '&', idx_59, to)) > -1 && idx_38 < to ? indexOf(cb, ';', idx_38 + 1, Math.min(idx_38 + 9, to)) : -1;
        } while (idx_38 > -1 && idx_59 > idx_38 && idx_59 <= to);

        if (from < to) {
            System.arraycopy(cb, from, buf, offset, n = to - from);
            offset += n;
        }

        return new String(buf, 0, offset);
    }

    private static int indexOf(char[] cb, char c, int from, int to) {
        for (; from < to && cb[from] != c; from++);
        return from < to ? from : -1;
    }

    static int toInt(char[] cb, int a, int b) {
        int n = 0;
        boolean negative = cb[a] == '-';
        for (int i = negative ? a + 1 : a; b > i; n = n * 10 + cb[i++] - '0');
        return negative ? -n : n;
    }

    // Hex value
    static int toIntH(char[] cb, int a, int b) {
        int n = 0;
        for (int c; b > a; n = n * 16 + ((c = cb[a++]) <= '9' ? c - '0' : (c >= 'a' ? c - 32 : c) - '7'));
        return n;
    }

    static int toChars(int codePoint, char[] dst, int i) {
        int n;
        if (isBmpCodePoint(codePoint)) {
            dst[i] = (char) codePoint;
            n = 1;
        } else if (isValidCodePoint(codePoint)) {
            dst[i + 1] = lowSurrogate(codePoint);
            dst[i] = highSurrogate(codePoint);
            n = 2;
        } else {
           dst[i] = MALFORMED_CHAR; // Illegal value �
           n = 1;
        }
        return n;
    }

    /**
     * close stream and free space
     */
    @Override
    public void close() throws IOException {
        if (reader != null) {
            // Debug hit rate
            LOGGER.debug("Count:{}, Forward:{}, Backward:{}, SST:{}, Hot:{}, Tester:{Resize:{}, Size:{}}", max
                , total_forward, total_backward, total_sst, total_hot
                , tester != null ? tester.analysis() : 0, tester != null ? tester.size() : 0);
            reader.close();
        }
        cb = null;
        forward = null;
        backward = null;
        if (tester != null) {
            tester = null;
        }
        if (sst != null) {
            sst.close();
        }
    }

    @Override
    public String toString() {
        return "Count: " + (total <= 0 ? max : total) + "，UniqueCount: " + max;
    }

}

interface Tester {

    /**
     * Test if a string needs to be cached
     *
     * @param i the string index in {@link IndexSharedStringTable}
     * @return true if the string should be cached
     */
    boolean test(int i);

    /**
     * Returns the limit index of {@link Tester}
     *
     * @return limit index
     */
    int limit();

    /**
     * Returns the block size of {@link Tester}
     *
     * @return the mark array length
     */
    int size();

    int analysis();

    class BinaryTester implements Tester {
        private int start;
        private int limit;
        private final int initial_size;
        private final long[] marks;
//        private static final int LIMIT = (1 << 25) - 1;

        private int total_resize; // For debug

        BinaryTester(int expectedInsertions) {
            marks = new long[initial_size = ((expectedInsertions - 1) >> 6) + 1];
            limit = (initial_size << 6) - 1;
        }

        @Override
        public boolean test(int i) {
            if (i < start) return true;
            // Check bound of bit-set
            if (i > limit && !resize(i)) return false;
            i = i - start;
            int n = i >> 6, m = i - (n << 6);
            boolean a = ((marks[n] >> (63 - m)) & 1) == 1;
            marks[n] |= 1L << (63 - m);
            return a;
        }

        @Override
        public int limit() {
            return limit;
        }

        @Override
        public int size() {
            return marks.length;
        }

        @Override
        public int analysis() {
            return total_resize;
        }

        private boolean resize(int i) {
            total_resize++;
            int ii = 0, n = marks.length, l = ((i - start) >> 6) + 1;

            for (; ii < n && marks[ii] == -1; ii++) ;
            if (l - ii > initial_size)
                for (; ii < n && (Long.bitCount(marks[ii]) > 48 || marks[ii] == 0); ii++) ;

            if (l - ii <= initial_size) {
                // Clean old mark
                int j = n - ii;
                if (j > 0) System.arraycopy(marks, ii, marks, 0, j);

                for (; j < n; marks[j++] = 0L) ;
                start += (ii << 6);
                limit = (marks.length << 6) + start - 1;
                return true;
            }
            return false;
        }
    }
}
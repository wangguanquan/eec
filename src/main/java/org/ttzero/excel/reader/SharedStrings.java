/*
 * Copyright (c) 2017-2018, guanquan.wang@yandex.com All Rights Reserved.
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
import org.ttzero.excel.entity.ExcelWriteException;

import java.io.BufferedReader;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;

import static org.ttzero.excel.util.FileUtil.exists;
import static org.ttzero.excel.util.StringUtil.EMPTY;

/**
 * Read sharedString data
 * <p>
 * For large files, it is impossible to load all data into the
 * memory and get it by index. The current practice is to partition
 * and divide the data into three areas: forward, backward, and hot.
 * The newly read data is placed in the forward area, if not found
 * in forward area, go to the backward area to find, finally go to
 * the hot area to find it. If not found in the three areas, press
 * the ward will be re-load in to the forward area. The original
 * forwarding area data is copied to the backward area. The blocks
 * loaded twice will be marked, the marked blocks will be placed in
 * the hot area when they are repeatedly read, and the hot area will
 * be eliminated by the LRU page replacement algorithm.
 *
 * @author guanquan.wang at 2018-09-27 14:28
 */
public class SharedStrings implements AutoCloseable {
    private final Logger LOGGER = LoggerFactory.getLogger(getClass());
    private Path sstPath;

    /**
     * Constructs a SharedStrings containing the elements of the
     * specified data array
     *
     * @param data the shared strings
     */
    SharedStrings(String[] data) {
        max = data.length;
        offset_forward = 0;
        status = 1;
        if (max <= page) {
            forward = new String[max];
            System.arraycopy(data, offset_forward, forward, 0, max);
            limit_forward = max;
        } else {
            if (max > page << 1) {
                page = max >> 1;
            }
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
     * @param sstPath   the xml file path
     * @param cacheSize the number of word per load
     * @param hotSize   the number of high frequency word
     */
    SharedStrings(Path sstPath, int cacheSize, int hotSize) {
        this.sstPath = sstPath;
        if (cacheSize > 0) {
            this.page = cacheSize;
        }
        this.hotSize = hotSize;
    }

    /**
     * Constructs a SharedStrings with a {@link IndexSharedStringTable}
     *
     * @param sst {@link IndexSharedStringTable}
     * @param cacheSize the number of word per load
     * @param hotSize   the number of high frequency word
     */
    SharedStrings(IndexSharedStringTable sst, int cacheSize, int hotSize) throws IOException {
        this.sst = sst;
        max = sst.size();
        if (cacheSize > 0) {
            this.page = cacheSize;
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
    private int page = 512;
    /**
     * The word total
     */
    private int max = -1, vt = 0, offsetM = 0;
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
    private BufferedReader reader;
    /**
     * Buffered
     */
    private char[] cb;
    /**
     * offset of cb[]
     */
    private int offset;
    /**
     * escape buffer
     */
    private StringBuilder escapeBuf;
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
     * Load the sharedString.xml file and instance word cache
     *
     * @return the {@code SharedStrings}
     * @throws IOException if io error occur
     */
    SharedStrings load() throws IOException {
        if (!exists(sstPath)) {
            max = 0;
            return this;
        }

        // Get unique count
        max = uniqueCount();
        LOGGER.debug("Size of SharedString: {}", max);
        //
        init();
        escapeBuf = new StringBuilder();
        return this;
    }

    /* */
    private void init() throws IOException {
        status = 1;
        // Unknown size or greater than 512
        if (max < 0 || max > page << 1) {
            status <<= 2;
            forward = new String[page];
            backward = new String[page];

            // Cache 8KB binary, it will store 1^16 strings.
            tester = new Tester.FixBinaryTester(max > 1 << 16 ? 1 << 16 : max);

            if (hotSize > 0) hot = FixSizeLRUCache.create(hotSize);
            else hot = FixSizeLRUCache.create();
            // Instance the SharedStringTable
            if (sst == null) {
                sst = new IndexSharedStringTable();
                sst.setShortSectorSize(9);
            }
        }
        else if (max > page) {
            status <<= 1;
            forward = new String[page];
            backward = new String[page];
        } else {
            forward = new String[max];
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
        reader = Files.newBufferedReader(sstPath);
        cb = new char[1 << 12];
        offset = 0;
        offset = reader.read(cb);

        int i = 0, len = offset - 4;
        for (; i < len && (cb[i] != '<' || cb[i + 1] != 's' || cb[i + 2] != 'i' || cb[i + 3] != '>'); i++) ;
        if (i == len) return 0; // Empty

        String line = new String(cb, 0, i);
        // Microsoft Excel
        String uniqueCount = " uniqueCount=";
        int index = line.indexOf(uniqueCount)
            , end = index > 0 ? line.indexOf('"'
            , index += (uniqueCount.length() + 1)) : -1;
        if (end > 0) {
            off = Integer.parseInt(line.substring(index, end));
            // WPS
        } else {
            String count = " count=";
            index = line.indexOf(count);
            end = index > 0 ? line.indexOf('"', index += (count.length() + 1)) : -1;
            if (end > 0) {
                off = Integer.parseInt(line.substring(index, end));
            }
        }

        vt = i + 4;
        System.arraycopy(cb, vt, cb, 0, offset -= vt);

        return off;
    }

    /**
     * Getting the strings value by index
     *
     * @param index the index of SharedStringTable
     * @return string
     */
    public String get(int index) {
        checkBound(index);
        total++;

        // Load first
        if (offset_forward == -1) {
            offset_forward = index / page * page;

            if (vt < 0) vt = 0;

            loadXml();
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

            System.arraycopy(forward, 0, backward, 0, limit_forward);
            offset_backward = offset_forward;
            limit_backward = limit_forward;
            // reload data
            offset_forward = index / page * page;
            forward[0] = null;
            if (status == 4 && index < sst.size()) {
                try {
                    // Load from SharedStringTable
                    limit_forward = sst.get(offset_forward, forward);
                } catch (IOException e) {
                    throw new ExcelWriteException(e);
                }
                total_sst++;
            } else {
                loadXml();
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

    // Check the current index if out of bound
    private void checkBound(int index) {
        if (index < 0 || max > -1 && max <= index) {
            throw new IndexOutOfBoundsException("Index: " + index + ", Size: " + max);
        }
    }

    // Check the current index has been loaded twice
    private boolean test(int index) {
        return status == 4 && tester.test(index);
    }

    /**
     * Load string record from xml
     */
    private void loadXml() {
        int index = offset_forward / page;
        try {
            // Read xml file string value into IndexSharedStringTable
            for (int n = index - offsetM; n-- >= 0; ) {
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
    private int readData() throws IOException {
        // Read forward area data
        int n = 0, length, nChar;
        for (; (length = reader.read(cb, offset, cb.length - offset)) > 0 || offset > 0; ) {
            length += offset;
            nChar = offset &= 0;
            int len0 = length - 3, len1 = len0 - 1;
            int[] t = findT(cb, nChar, length, len0, len1, n);

            nChar = t[0];
            n = t[1];

            limit_forward = n;

            if (nChar < length) {
                System.arraycopy(cb, nChar, cb, 0, offset = length - nChar);
            }

            // A page
            if (n == page) {
                ++offsetM;
                break;
            } else if (length < cb.length && nChar == length - 6) { // EOF '</sst>'
                if (max == -1) { // Reset totals when unknown size
                    max = offsetM * page + n;
                }
                ++offsetM; // out of index range
                break;
            }
        }
        return n; // Returns the word count
    }

    private int[] findT(char[] cb, int nChar, int length, int len0, int len1, int n) throws IOException {
        int cursor;
        for (; nChar < length && n < page; ) {
            cursor = nChar;
            int[] subT = subT(cb, nChar, len0, len1);
            int a = subT[0];
            if (a == -1) break;
            nChar = subT[1];

            String tmp = unescape(escapeBuf, cb, a, nChar);

             // Skip the end tag of 't'
            nChar += 4;

            // Test the next tag
            for (; nChar < len1 && (cb[nChar] != '<' || cb[nChar + 1] != '/'); ++nChar);

            // End of <si>
            if (nChar < len1 && cb[nChar + 2] == 's' && cb[nChar + 3] == 'i' && cb[nChar + 4] == '>') {
                forward[n++] = tmp;
                if (status == 4) sst.push(forward[n - 1]);
                nChar += 5;
            } else {
                StringBuilder buf = new StringBuilder(tmp);
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
                // Loop and join
                for (; ; ) {
                    subT = subT(cb, nChar, end, end);
                    a = subT[0];
                    if (a == -1) break;
                    nChar = subT[1];

                    buf.append(unescape(escapeBuf, cb, a, nChar));
                    nChar += 4;
                }
                forward[n++] = buf.toString();
                if (status == 4) sst.push(forward[n - 1]);
            }

            // An integral page records
            if (n == page) break;
        }
        // DEBUG the last character
//        LOGGER.info("---------{}---------", new String(cb, nChar, length - nChar));
        return new int[] { nChar, n };
    }

    private int[] subT(char[] cb, int nChar, int len0, int len1) {

        for (; nChar < len0 && (cb[nChar] != '<' || cb[nChar + 1] != 't'
            || cb[nChar + 2] != '>' && cb[nChar + 2] != ' '); ++nChar)
            ;
        if (nChar >= len0) return new int[]{ -1 }; // Not found
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

    /**
     * unescape
     */
    static String unescape(StringBuilder escapeBuf, char[] cb, int from, int to) {
        if (from == to) return EMPTY;
        int idx_38 = indexOf(cb, '&', from)
            , idx_59 = idx_38 > -1 && idx_38 < to ? indexOf(cb, ';', idx_38 + 1) : -1;

        if (idx_38 <= 0 || idx_38 >= idx_59 || idx_59 > to) {
            return new String(cb, from, to - from);
        }
        escapeBuf.delete(0, escapeBuf.length());
        do {
            escapeBuf.append(cb, from, idx_38 - from);
            // ASCII值
            if (cb[idx_38 + 1] == '#') {
                int n = toInt(cb, idx_38 + 2, idx_59);
                // byte range
//                if (n < 0 || n > 127) {
//                    LOGGER.warn("Unknown escape [{}]", new String(cb, idx_38, idx_59 - idx_38 + 1));
//                }
                // Unicode char
                escapeBuf.append((char) n);
            }
            // 转义字符
            else {
                String name = new String(cb, idx_38 + 1, idx_59 - idx_38 - 1);
                switch (name) {
                    case "lt":
                        escapeBuf.append('<');
                        break;
                    case "gt":
                        escapeBuf.append('>');
                        break;
                    case "amp":
                        escapeBuf.append('&');
                        break;
                    case "quot":
                        escapeBuf.append('"');
                        break;
                    case "nbsp":
                        escapeBuf.append(' ');
                        break;
                    default: // Unknown escape
//                        LOGGER.warn("Unknown escape [&{}]", name);
                        escapeBuf.append(cb, idx_38, idx_59 - idx_38 + 1);
                }
            }
            from = ++idx_59;
            idx_59 = (idx_38 = indexOf(cb, '&', idx_59)) > -1 && idx_38 < to ? indexOf(cb, ';', idx_38 + 1) : -1;
        } while (idx_38 > -1 && idx_59 > idx_38 && idx_59 <= to);

        if (from < to) {
            escapeBuf.append(cb, from, to - from);
        }
        return escapeBuf.toString();
    }

    private static int indexOf(char[] cb, char c, int from) {
        for (; from < cb.length; from++) {
            if (cb[from] == c) return from;
        }
        return -1;
    }

    static int toInt(char[] cb, int a, int b) {
        boolean _n;
        if (_n = cb[a] == '-') a++;
        int n = cb[a++] - '0';
        for (; b > a; ) {
            n = n * 10 + cb[a++] - '0';
        }
        return _n ? -n : n;
    }

    /**
     * close stream and free space
     */
    @Override
    public void close() throws IOException {
        if (reader != null) {
            // Debug hit rate
            LOGGER.debug("Count: {}， uniqueCount: {}， Repetition rate: {}", total, max
                , (total > 0 ? (total - max) * 100.0/ total : 0) + "%");
            LOGGER.debug("Forward: {}, Backward: {}, SST: {}, Hot: {}, Tester: {Resize: {}, Size: {}}"
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
        escapeBuf = null;
        if (sst != null) {
            sst.close();
        }
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

    class FixBinaryTester implements Tester {
        private int start, limit, initial_size;
        private long[] marks;
//        private static final int LIMIT = (1 << 25) - 1;

        private int total_resize; // For debug

        FixBinaryTester(int expectedInsertions) {
            marks = new long[initial_size = ((expectedInsertions - 1) >> 6) + 1];
            limit = (initial_size << 6) - 1;
        }

        @Override
        public boolean test(int i) {
            if (i < start) return true;
            // Check bound of bit-set
            if (i > limit) resize(i);
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

        private void resize(int i) {
            total_resize++;
            int ii = 0, n = marks.length, l = ((i - start) >> 6) + 1;

            for (; ii < n && marks[ii] == -1; ii++) ;

            if (l - ii <= initial_size) {
                // Clean old mark
                if (ii > 0) {
                    int j = 0;
                    for (int m = ii; m < n; marks[j++] = marks[m++]) ;
                    for (; j < n; marks[j++] &= 0) ;
                    start += (ii << 6);
                }
//                else {
//                    marks = Arrays.copyOf(marks, l + (l >> 1));
//                }
            } else {
                // TODO Limit Tester length to prevent infinite expansion
                long[] newMarks = new long[(l - ii) + ((l - ii) >> 1)];
                System.arraycopy(marks, ii, newMarks, 0, marks.length - ii);
                marks = newMarks;
                start += (ii << 6);
            }
            limit = (marks.length << 6) + start - 1;
        }
    }
}
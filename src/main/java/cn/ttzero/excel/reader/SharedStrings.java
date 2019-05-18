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

package cn.ttzero.excel.reader;

import cn.ttzero.excel.entity.ExcelWriteException;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import java.io.BufferedReader;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.*;

/**
 * Read sharedString data
 *
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
 * Create by guanquan.wang at 2018-09-27 14:28
 */
public class SharedStrings implements AutoCloseable {
    private Logger logger = LogManager.getLogger(getClass());
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
        if (max <= page) {
            forward = new String[max];
            System.arraycopy(data, offset_forward, forward, 0, max);
            limit_forward = max;
        } else {
            if (max > page << 1) {
                page = max >> 1;
            }
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
     * @param sstPath the xml file path
     * @param cacheSize the number of word per load
     * @param hotSize the number of high frequency word
     */
    SharedStrings(Path sstPath, int cacheSize, int hotSize) {
        this.sstPath = sstPath;
        if (cacheSize > 0) {
            this.page = cacheSize;
        }
        this.hotSize = hotSize;
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
     * Count the number of visits to each segment
     */
    private Map<Integer, Integer> count_area = null;
    /**
     * High frequency word
     */
    private Hot<Integer, String> hot;
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

    private IndexSharedStringTable sst;
    // For debug
    private int total, total_forward, total_backward, total_hot;

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
        if (Files.exists(sstPath)) {
            // Get unique count
            max = uniqueCount();
            logger.debug("Size of SharedString: {}", max);
            // Unknown size or big than page
            int default_cap = 10;
            if (max < 0 || max > page) {
                forward = new String[page];
                backward = new String[page];

                if (max > 0 && max / page + 1 > default_cap)
                    default_cap = max / page + 1;
                count_area = new HashMap<>(default_cap);

                if (hotSize > 0) hot = new Hot<>(hotSize);
                else hot = new Hot<>();
            } else {
                forward = new String[max];
            }
        } else {
            max = 0;
        }
        escapeBuf = new StringBuilder();
        // Instance the SharedStringTable
        // FIXME split size
        sst = new IndexSharedStringTable();
        return this;
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
        cb = new char[2048];
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
            test(index);
        }

        String value;
        // Find in forward
        if (forwardRange(index)) {
            value = forward[index - offset_forward];
            total_forward++;
            return value;
        }

        // Find in backward
        if (backwardRange(index)) {
            value = backward[index - offset_backward];
            total_backward++;
            return value;
        }

        // Find in hot cache
        value = hot.get(index);

        // Can't find in memory cache
        if (value == null) {
            System.arraycopy(forward, 0, backward, 0, limit_forward);
            offset_backward = offset_forward;
            limit_backward = limit_forward;
            // reload data
            offset_forward = index / page * page;
            forward[0] = null;
            if (index < sst.size()) {
                try {
                    // Load from SharedStringTable
                    limit_forward = sst.get(offset_forward, forward);
                } catch (IOException e) {
                    throw new ExcelWriteException(e);
                }
            } else {
                loadXml();
            }
            if (forward[0] == null) {
                throw new IndexOutOfBoundsException("Index: " + index + ", Size: " + max);
            }
            value = forward[index - offset_forward];
            if (test(index)) {
                logger.debug("put hot {}", index);
                hot.push(index, value);
            }
            total_forward++;
        } else {
            total_hot++;
        }

        return value;
    }

    private boolean forwardRange(int index) {
        return offset_forward >= 0 && offset_forward <= index
            && offset_forward + limit_forward > index;
    }

    private boolean backwardRange(int index) {
        return offset_backward >= 0 && offset_backward <= index
            && offset_backward + limit_backward > index;
    }

    private void checkBound(int index) {
        if (index < 0 || max > -1 && max <= index) {
            throw new IndexOutOfBoundsException("Index: " + index + ", Size: " + max);
        }
    }

    /**
     * LRU2
     *
     * @param index the string value index
     * @return
     */
    private boolean test(int index) {
        if (max < page) return false;
        int idx = index / page;
        int n = count_area.getOrDefault(idx, 0) + 1;
        count_area.put(idx, n);
        return n > 1;
    }

    /**
     * Read or Load xml
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
     * @return word count
     * @throws IOException if I/O error occur
     */
    private int readData() throws IOException {
        // Read forward area data
        int n = 0, length, nChar;
        while ((length = reader.read(cb, offset, cb.length - offset)) > 0 || offset > 0) {
            length += offset;
            nChar = offset &= 0;
            int cursor, len0 = length - 3, len1 = len0 - 1;
            int[] t = findT(cb, nChar, length, len0, len1, n);

            nChar = t[0];
            n = t[1];
            cursor = t[2];

            limit_forward = n;

            if (cursor < length) {
                System.arraycopy(cb, cursor, cb, 0, offset = length - cursor);
            }

            // A page
            if (n == page) {
                ++offsetM;
                break;
            } else if (length < cb.length && nChar == len0) { // EOF
                if (max == -1) { // Reset totals when unknown size
                    max = offsetM * page + n;
                }
                ++offsetM; // out of index range
                break;
            }
        }
        return n; // Return word count
    }

    private int[] findT(char[] cb, int nChar, int length, int len0, int len1, int n) throws IOException {
        int cursor = 0;
        for (; nChar < length && n < page; ) {
            for (; nChar < len0 && (cb[nChar] != '<' || cb[nChar + 1] != 't'
                || cb[nChar + 2] != '>' && cb[nChar + 2] != ' '); ++nChar)
                ;
            if (nChar >= len0) break; // Not found
            cursor = nChar;
            int a = nChar += 3;
            if (cb[nChar - 1] == ' ') { // space="preserve"
                for (; nChar < len0 && cb[nChar++] != '>'; ) ;
                if (nChar >= len0) break; // Not found
                cursor = nChar;
                a = nChar;
            }
            for (; nChar < len1 && (cb[nChar] != '<' || cb[nChar + 1] != '/'
                || cb[nChar + 2] != 't' || cb[nChar + 3] != '>'); ++nChar)
                ;
            if (nChar >= len1) break; // Not found
            forward[n++] = nChar > a ? unescape(escapeBuf, cb, a, nChar) : null;
            sst.push(forward[n - 1]);
            nChar += 4;
            cursor = nChar;
        }
        return new int[]{nChar, n, cursor};
    }

    /**
     * unescape
     */
    static String unescape(StringBuilder escapeBuf, char[] cb, int from, int to) {
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
//                    logger.warn("Unknown escape [{}]", new String(cb, idx_38, idx_59 - idx_38 + 1));
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
//                        logger.warn("Unknown escape [&{}]", name);
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
            logger.debug("total: {}, forward: {}, backward: {}, hot: {}", total, total_forward, total_backward, total_hot);
            reader.close();
        }
        cb = null;
        forward = null;
        backward = null;
        if (count_area != null) {
            count_area.clear();
            count_area = null;
        }
        escapeBuf = null;
        if (sst != null) {
            sst.close();
        }
    }

}

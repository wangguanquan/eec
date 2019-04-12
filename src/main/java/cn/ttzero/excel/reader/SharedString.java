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

import cn.ttzero.excel.tmap.TIntIntHashMap;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import java.io.BufferedReader;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.*;

/**
 * 读取sharedString数据
 * 对于大文件来说不可能全部加载到内存然后根据下标直接取值，
 * 目前做法是进行分区，将数据分为eden, old, hot三个区域，
 * 新读取的数据放入eden区，如果eden区查找不到则去old区查找，最后到hot区查找
 * 如果都没有则按下标重新读取该区块到eden区，原eden区数据复制到old区。
 * 加载两次的区块将被标记，被标记的区块有重复读取时被放入hot区，
 * hot区采用LRU页面置换算法进行淘汰。
 *
 * Create by guanquan.wang at 2018-09-27 14:28
 */
public class SharedString {
    private Logger logger = LogManager.getLogger(getClass());
    private Path sstPath;

    SharedString(String[] data) {
        max = data.length;
        offset_eden = 0;
        if (max <= page) {
            eden = new String[max];
            System.arraycopy(data, offset_eden, eden, 0, max);
            limit_eden = max;
        } else {
            if (max > page << 1) {
                page = max >> 1;
            }
            eden = new String[page];
            limit_eden = page;
            System.arraycopy(data, offset_eden, eden, 0, limit_eden);
            offset_old = page;
            limit_old = max - page;
            old = new String[limit_old];
            System.arraycopy(data, offset_old, old, 0, limit_old);
        }
    }

    SharedString(Path sstPath, int cacheSize, int hotSize) {
        this.sstPath = sstPath;
        if (cacheSize > 0) {
            this.page = cacheSize;
        }
        this.hotSize = hotSize;
    }

    /** 新加载数据放入此区域 */
    private String[] eden;
    /** eden区未命中时将数据复制到此区域 */
    private String[] old;
    /** 每次加载的数量 */
    private int page = 512;
    /** 整个文件有多少个字符串 */
    private int max = -1, vt = 0, offsetR = 0, offsetM = 0;
    /** 新生代offset */
    private int offset_eden = -1;
    /** 老年代offset */
    private int offset_old = -1;
    /** 新生代limit */
    private int limit_eden;
    /** 老年代limit */
    private int limit_old;
    /** 统计各段被访问次数 */
    private TIntIntHashMap count_area = null;
    /** hot world */
    private Hot hot;
    /** size of hot */
    private int hotSize;
    /** 记录各段的起始位置 */
    private Map<Integer, Long> index_area = null;

    /** Main reader */
    private BufferedReader reader;
    /** Buffered */
    private char[] cb;
    /** offset of cb[] */
    private int offset;
    /** current skip of Main reader */
    private long skipR;
    /** escape buffer */
    private StringBuilder escapeBuf;
    // For debug
    private int total, total_eden, total_old, total_hot;

    /**
     * @return the shared string unique count
     * -1 if unknown size
     */
    int size() {
        return max;
    }

    SharedString load() throws IOException {
        if (Files.exists(sstPath)) {
            // Get unique count
            max = uniqueCount();
            logger.debug("Size of SharedString: {}", max);
            // Unknown size or big than page
            int default_cap = 10;
            if (max < 0 || max > page) {
                eden = new String[page];
                old = new String[page];

                if (max > 0 && max / page + 1 > default_cap) default_cap = max / page + 1;
                count_area =  new TIntIntHashMap(default_cap);

                if (hotSize > 0) hot = new Hot(hotSize);
                else hot = new Hot();
            }
            else {
                eden = new String[max];
            }
            index_area = new HashMap<>(default_cap);
            index_area.put(0, (long) vt);
        } else {
            max = 0;
        }
        return this;
    }

    private int uniqueCount() throws IOException {
        int off = -1;
        reader = Files.newBufferedReader(sstPath);
        cb = new char[2048];
        offset = 0;
        offset = reader.read(cb);

        int i = 0, len = offset - 4;
        for (; i < len && (cb[i] != '<' || cb[i+1] != 's' || cb[i+2] != 'i' || cb[i+3] != '>'); i++);
        if (i == len) return  0; // Empty

        String line = new String(cb, 0, i);
        // Microsoft Excel
        String uniqueCount = " uniqueCount=";
        int index = line.indexOf(uniqueCount), end = index > 0 ? line.indexOf('"', index += (uniqueCount.length() + 1)) : -1;
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
        skipR = vt;

        return off;
    }

    /**
     * 根据下标获取字符串
     * @param index of sharedString
     * @return string
     */
    String get(int index) {
        total++;
        if (index < 0 || max > -1 && max <= index) {
            throw new IndexOutOfBoundsException("Index: "+index+", Size: "+max);
        }

        // Load first
        if (offset_eden == -1) {
            offset_eden = index / page * page;

            if (vt < 0) vt = 0;

            loadXml();
            test(index);
        }

        String value;
        // Find in eden
        if (edenRange(index)) {
            value = eden[index - offset_eden];
            total_eden++;
            return value;
        }

        // Find in old
        if (oldRange(index)) {
            value = old[index - offset_old];
            total_old++;
            return value;
        }

        // Find in hot cache
        value = hot.get(index);

        // Can't find in memory cache
        if (value == null) {
            System.arraycopy(eden, 0, old, 0, limit_eden);
            offset_old = offset_eden;
            limit_old = limit_eden;
            // reload data
            offset_eden = index / page * page;
            eden[0] = null;
            loadXml();
            if (eden[0] == null) {
                throw new IndexOutOfBoundsException("Index: "+index+", Size: "+max);
            }
            value = eden[index - offset_eden];
            if (test(index)) {
                logger.debug("put hot {}", index);
                hot.push(index, value);
            }
            total_eden++;
        } else {
            total_hot++;
        }

        return value;
    }

    private boolean edenRange(int index) {
        return offset_eden >= 0 && offset_eden <= index && offset_eden + limit_eden > index;
    }

    private boolean oldRange(int index) {
        return offset_old >= 0 && offset_old <= index && offset_old + limit_old > index;
    }

    /**
     * LRU2
     * @param index
     * @return
     */
    private boolean test(int index) {
        if (max < page) return false;
        int n = count_area.incrementGet(index / page);
        return n > 1;
    }

    /**
     * Read or Load xml
     */
    private void loadXml() {
        int index = offset_eden / page;
        try {
            /*
             * 从cache流中读取接下来的N个数据
             * 如果Excel保持从上到下从左到右的写入顺序
             * 那么此分支命中率极高，当此分支命中时就会直接从已知句柄中直接读取N个字符
             * 减少读取文件的次数，同时也减少读取相同区块的次数，对于大文件读取效率非常高。
             */
            if (index == offsetM) {
                logger.debug("---------------Read xml area {}---------------", index);
                readData();
            }
            /*
             * 从文件读取
             * 文件读取时会跳过已知区块以增加速度（虽然java.io中skip方法并非跳过读取而是读取后跳过）
             * 如果先读取文件尾部的数据那么在skip的时候会记录每个区块的起始位置，
             * 接下来的读取将根据此预存的记录进行skip
             */
            else {
                logger.debug("---------------Load xml area {}---------------", index);
                loadXml(index);
            }
        } catch (IOException e) {
            throw new ExcelReadException(e);
        }
    }

    /**
     * 分段加载数据
     */
    private void loadXml(int index) throws IOException {
        try (BufferedReader br = Files.newBufferedReader(sstPath)) {
            char[] cb = new char[this.cb.length];
            int nChar, length, n = 0, offset = 0;
            long skip = index_area.getOrDefault(index, 0L), skipR = skip;

            // 跳过N个区域读取时预先记录跳过的区域起始位置
            if (skip <= 0L) {
                if (index_area.isEmpty()) {
                    br.skip(skipR = vt);
                } else {
                    br.skip(skipR = index_area.get(offsetR));
                }

                while (offsetR < index && (length = br.read(cb, offset, cb.length - offset)) > 0) {
                    nChar = 0;
                    length += offset;
                    int cursor = 0;
                    for (int len = length - 4; nChar < len && n < page; nChar++) {
                        if (cb[nChar] == '<' && cb[nChar + 1] == '/' && cb[nChar + 2] == 't' && cb[nChar + 3] == '>') {
                            n++;
                            cursor = nChar += 4;
                        } else continue;
                        // a page
                        if (n == page) {
                            n = 0;
                            index_area.put(++offsetR, skipR + cursor);
                            if (offsetR >= index) break;
                        }
                    }
                    skipR += cursor;
                    System.arraycopy(cb, cursor, cb, 0, offset = length - cursor);
                }
                // 下一个相邻区域或者已读区域
            } else {
                br.skip(skipR); // 跳过前方N个字符从当前位置直接读取
            }

            // Read eden area data
            n = 0;
            while ((length = br.read(cb, offset, cb.length - offset)) > 0 || offset > 0) {
                length += offset;
                nChar = offset &= 0;
                int cursor, len0 = length - 3, len1 = len0 - 1;
                int[] t = findT(cb, nChar, length, len0, len1, n);

                nChar = t[0];
                n = t[1];
                cursor = t[2];

                limit_eden = n;
                skipR += cursor;

                // A page
                if (n == page) {
                    // Save next start index
                    if (offsetR <= index) {
                        index_area.put(++offsetR, skipR);
                    }
                    break;
                } else if (length < cb.length && nChar == len0) { // EOF
                    if (max == -1) { // Reset totals when unknown size
                        max = offsetR * page + n;
                    }
                    break;
                }

                if (cursor < length) {
                    System.arraycopy(cb, cursor, cb, 0, offset = length - cursor);
                }
            }
        }
    }

    /**
     * Read data from main reader
     * forward only
     * @return word count
     * @throws IOException -
     */
    private int readData() throws IOException {
        // Read eden area data
        int n = 0, length, nChar;
        while ((length = reader.read(cb, offset, cb.length - offset)) > 0 || offset > 0) {
            length += offset;
            nChar = offset &= 0;
            int cursor, len0 = length - 3, len1 = len0 - 1;
            int[] t = findT(cb, nChar, length, len0, len1, n);

            nChar = t[0];
            n = t[1];
            cursor = t[2];

            limit_eden = n;
            skipR += cursor;

            if (cursor < length) {
                System.arraycopy(cb, cursor, cb, 0, offset = length - cursor);
            }

            // A page
            if (n == page) {
                // Save next start index
                index_area.put(++offsetM, skipR);
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

    private int[] findT(char[] cb, int nChar, int length, int len0, int len1, int n) {
        int cursor = 0;
        for ( ; nChar < length && n < page; ) {
            for (; nChar < len0 && (cb[nChar] != '<' || cb[nChar + 1] != 't' || cb[nChar + 2] != '>' && cb[nChar + 2] != ' '); ++nChar);
            if (nChar >= len0) break; // Not found
            cursor = nChar;
            int a = nChar += 3;
            if (cb[nChar - 1] == ' ') { // space="preserve"
                for (; nChar < len0 && cb[nChar++] != '>'; );
                if (nChar >= len0) break; // Not found
                cursor = nChar;
                a = nChar;
            }
            for (; nChar < len1 && (cb[nChar] != '<' || cb[nChar + 1] != '/' || cb[nChar + 2] != 't' || cb[nChar + 3] != '>'); ++nChar);
            if (nChar >= len1) break; // Not found
            eden[n++] = nChar > a ? unescape(cb, a, nChar) : null;
            nChar += 4;
            cursor = nChar;
        }
        return new int[]{ nChar, n, cursor };
    }

    /**
     * 反转义
     */
    String unescape(char[] cb, int from, int to) {
        int idx_38 = indexOf(cb, '&', from), idx_59 = idx_38 > -1 && idx_38 < to ? indexOf(cb, ';', idx_38 + 1) : -1;

        if (idx_38 <= 0 || idx_38 >= idx_59 || idx_59 > to) {
            return new String(cb, from, to - from);
        }
        if (escapeBuf != null) {
            escapeBuf.delete(0, escapeBuf.length());
        } else {
            escapeBuf = new StringBuilder(to - from < 10 ? 10 : to - from);
        }
        do {
            escapeBuf.append(cb, from, idx_38 - from);
            // ASCII值
            if (cb[idx_38 + 1] == '#') {
                int n = toInt(cb, idx_38+2, idx_59);
                // byte range
//                if (n < 0 || n > 127) {
//                    logger.warn("Unknown escape [{}]", new String(cb, idx_38, idx_59 - idx_38 + 1));
//                }
                // Unicode char
                escapeBuf.append((char) n);
            }
            // 转义字符
            else {
                String name = new String(cb, idx_38+1, idx_59 - idx_38 - 1);
                switch (name) {
                    case "lt": escapeBuf.append('<'); break;
                    case "gt": escapeBuf.append('>'); break;
                    case "amp": escapeBuf.append('&'); break;
                    case "quot": escapeBuf.append('"'); break;
                    case "nbsp": escapeBuf.append(' '); break;
                    default: // Unknown escape
                        logger.warn("Unknown escape [&{}]", name);
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

    private int indexOf(char[] cb, char c, int from) {
        for ( ; from < cb.length; from++) {
            if (cb[from] == c) return from;
        }
        return -1;
    }

    private int toInt(char[] cb, int a, int b) {
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
    void close() {
        if (reader != null) try {
            // Debug hit rate
            logger.debug("total: {}, eden: {}, old: {}, hot: {}", total, total_eden, total_old, total_hot);
            reader.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
        cb = null;
        eden = null;
        old = null;
        if (index_area != null) {
            index_area.clear();
            index_area = null;
        }
        if (count_area != null) {
            count_area.clear();
            count_area = null;
        }
        escapeBuf = null;
    }
}

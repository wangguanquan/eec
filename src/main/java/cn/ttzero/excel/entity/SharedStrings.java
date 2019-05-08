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

package cn.ttzero.excel.entity;

import cn.ttzero.excel.manager.Const;
import cn.ttzero.excel.annotation.TopNS;
import cn.ttzero.excel.reader.Hot;
import cn.ttzero.excel.util.ExtBufferedWriter;
import cn.ttzero.excel.util.FileUtil;
import cn.ttzero.excel.util.StringUtil;
import com.google.common.hash.BloomFilter;
import com.google.common.hash.Funnels;

import java.io.BufferedReader;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.ByteBuffer;
import java.nio.channels.FileChannel;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardOpenOption;
import java.util.Arrays;

import static cn.ttzero.excel.reader.SharedString.unescape;

/**
 * A workbook collects the strings of all text cells in a global list,
 * the Shared String Table. This table is located in the record SST in
 * the Workbook Globals Substream.
 *
 * SST saves characters and strings sequentially. When writing a string,
 * it first determines whether it exists. If it exists, returns the index
 * in the Table (zero base), otherwise add it in to the last element of
 * Table and returns the current subscript.
 * Introduced Google BloomFilter to increase filtering speed, the
 * BloomFilter estimates the amount of data to be 1 million, and the false
 * positive rate is 0.01%. When the number exceeds 1 million, it will not be
 * written to SST and will be converted to inline string.
 *
 * A hot zone is also designed internally to cache multiple occurrences,
 * the default size is 1024, and the LRU elimination algorithm is used.
 * If the cache misses, it will be read from in temp file and flushed to the
 * cache. In order to prevent too many time-consuming searches for only the
 * first 100,000 words, will converted to inline string if not found.
 *
 * Characters are handled differently. ASCII characters use the built-in array
 * cache subscript. The over 0x7F characters will be converted to strings and
 * searched using strings.
 *
 * Created by guanquan.wang on 2017/10/10.
 */
@TopNS(prefix = "", value = "sst", uri = Const.SCHEMA_MAIN)
public class SharedStrings implements Storageable, AutoCloseable {

    /**
     * The total word in workbook.
     */
    private int count;

//    /**
//     * The total unique word in workbook.
//     */
//    private int uniqueCount;

    /**
     * Import google BloomFilter to check keyword not exists
     */
    private BloomFilter<String> filter;

    /**
     * Cache ASCII value
     */
    private int[] ascii;

    /**
     * Cache the string which read twice and above
     * Use LRU elimination algorithm
     */
    private Hot<String, Integer> hot;

    /**
     * Storage into temp file on disk
     */
    private SharedStringTable sst;

    /**
     * the number of expected insertions to the constructed bloom
     */
    private int expectedInsertions = 1 << 20;

//    private StringBuilder buf;

    SharedStrings() {
        hot = new Hot<>(1 << 10);
        ascii = new int[1 << 7];
        // -1 means the keyword not exists
        Arrays.fill(ascii, -1);
        // Create a 2^20 expected insertions and 0.01% fpp bloom filter
        filter = BloomFilter.create(Funnels.stringFunnel(StandardCharsets.UTF_8), expectedInsertions, 0.0001);
        // The shared string table
        sst = new SharedStringTable();
    }

    private ThreadLocal<char[]> charCache = ThreadLocal.withInitial(() -> new char[1]);

    /**
     * Getting the character value index (zero base)
     *
     * @param c the character value
     * @return the index in ShareString
     */
    public int get(char c) throws IOException {
        // An ASCII keyword
        if (c < 128) {
            int n = ascii[c];
            // Not exists
            if (n == -1) {
                n = sst.push(c);
//                n = uniqueCount++;
                ascii[c] = n;
            }
            count++;
            return n;
        } else {
            // TODO write as character
            char[] cs = charCache.get();
            cs[0] = c;
            return get(new String(cs));
        }
    }

    /**
     * Getting the string value from cache (zero base)
     *
     * @param key the string value
     * @return index of the string in the SST
     * -1 if cache full, please write as 'inlineStr'
     */
    public int get(String key) throws IOException {
        count++;
        // The keyword not exists
        if (!filter.mightContain(key)) {
            // Add to bloom if not full
            if (sst.size() < expectedInsertions) {
                filter.put(key);
                return sst.push(key);
            } else return -1;
        }
        // Check the keyword exists in cache
        Integer n = hot.get(key);
        if (n == null) {
            // Find in temp file
            n = sst.find(key);
            // If not found in first 100,000 words
            // append last and cache it
            if (n < 0) {
                n = sst.push(key);
            }
            hot.push(key, n);
        }
        return n;
    }

//    private void add(String key) throws IOException {
//        writer.write("<si><t>");
//        writer.escapeWrite(key);
//        writer.write("</t></si>");
//    }
//
//    private void add(char c) throws IOException {
//        writer.write("<si><t>");
//        writer.escapeWrite(c);
//        writer.write("</t></si>");
//    }

    @Override
    public void writeTo(Path root) throws IOException {
        // TODO Close temp writer
//        FileUtil.close(writer);
        sst.close();

        if (!Files.exists(root)) {
            FileUtil.mkdir(root);
        }

        StringBuilder buf = new StringBuilder();
        TopNS topNS = getClass().getAnnotation(TopNS.class);
        if (topNS != null) {
            buf.append(Const.EXCEL_XML_DECLARATION);
            buf.append(Const.lineSeparator);
            buf.append("<").append(topNS.value()).append(" xmlns=\"").append(topNS.uri()[0]).append("\"")
                .append(" count=\"").append(count).append("\"")
                .append(" uniqueCount=\"").append(sst.size()).append("\">")
                .append(Const.lineSeparator);
        } else {
            buf.append("<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"")
                .append(count).append("\" uniqueCount=\"").append(sst.size()).append("\">")
                .append(Const.lineSeparator);
        }

        // The output path
        Path dist = root.resolve(StringUtil.lowFirstKey(getClass().getSimpleName() + Const.Suffix.XML));

        try (FileOutputStream fos = new FileOutputStream(dist.toFile());
             FileChannel channel = fos.getChannel()) {
            ByteBuffer buffer = ByteBuffer.allocate(1 << 9);
            buffer.put(buf.toString().getBytes(StandardCharsets.UTF_8));
            buffer.flip();
            channel.write(buffer);

            if (count > 0) {
                transfer(channel);
            }

            buffer.clear();
            buf.delete(0, buf.length());
            if (topNS != null) {
                buf.append("</").append(topNS.value()).append(">");
            } else {
                buf.append("</sst>");
            }
            buffer.put(buf.toString().getBytes(StandardCharsets.UTF_8));
            buffer.flip();
            channel.write(buffer);
        }
    }

    /**
     * Transfer temp data to dist path
     *
     * @param channel the dist file channel
     * @throws IOException if io error occur
     */
    private void transfer(FileChannel channel) throws IOException {
//        try (FileChannel tempChannel = FileChannel.open(temp, StandardOpenOption.READ)) {
//            tempChannel.transferTo(0, tempChannel.size(), channel);
//        }
        // TODO
    }
//
//    /**
//     * Found in temp file
//     * @param key the string value
//     * @return the index in sst file (zero base)
//     */
//    private int findFromFile(String key) throws IOException {
//        writer.flush();
//
//        buf = new StringBuilder();
//
//        try (BufferedReader reader = Files.newBufferedReader(temp, StandardCharsets.UTF_8)) {
//            char[] cb = new char[1 << 11];
//            int n, off = 0;
//            O o = new O();
//            o.cb = cb;
//            o.key = key;
//            while ((n = reader.read(cb, off, cb.length - off)) > 0) {
//                o.n = n + off;
//                findT(o);
//                // Get it
//                if (o.f) {
//                    return o.index;
//                    // search next block
//                } else if (o.off > 0) {
//                    System.arraycopy(cb, o.off, cb, 0, off = o.n - o.off);
//                    o.off = 0;
//                    // resize
//                } else {
//                    char[] bigCb = new char[cb.length << 1];
//                    System.arraycopy(cb, 0, bigCb, 0, o.n);
//                    o.cb = cb = bigCb;
//                    off = o.n;
//                }
//            }
//        }
//
//        return -1;
//    }
//
//    private void findT(O o) {
//        int sn = 7, ln = 9, nChar = sn;
//        for (; nChar < o.n; ) {
//            int next = next(o.cb, nChar, o.n);
//            // End of block
//            if (next >= o.n) {
//                break;
//            }
//            String v = unescape(buf, o.cb, nChar, next);
//            // Check values
//            o.f = v.equals(o.key);
//            if (o.f) {
//                break;
//            }
//            if (next + ln > o.n) {
//                o.off = nChar;
//                break;
//            }
//            // Not found in first 100,000 words
//            if (o.index++ >= find_limit) {
//                o.f = true;
//                o.index = -1;
//                break;
//            }
//
//            if (next + ln + sn > o.n) {
//                o.off = next + ln;
//                break;
//            }
//            nChar = next + ln + sn;
//        }
//    }
//
//    // found the end index of string value
//    private int next(char[] cb, int nChar, int n) {
//        for (; nChar < n && cb[nChar] != '<'; nChar++) ;
//        return nChar;
//    }

    @Override
    public void close() {
//        buf = null;
        filter = null;
        hot.clear();
        hot = null;
    }

    private static class O {
        private char[] cb;
        private int n;
        private int off;
        private String key;

        private int index;
        private boolean f;
    }
}

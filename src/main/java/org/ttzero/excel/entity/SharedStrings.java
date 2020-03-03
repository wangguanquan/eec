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

package org.ttzero.excel.entity;

import com.google.common.hash.BloomFilter;
import com.google.common.hash.Funnels;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.ttzero.excel.annotation.TopNS;
import org.ttzero.excel.manager.Const;
import org.ttzero.excel.reader.Cache;
import org.ttzero.excel.reader.FixSizeLRUCache;
import org.ttzero.excel.util.ExtBufferedWriter;
import org.ttzero.excel.util.FileUtil;
import org.ttzero.excel.util.StringUtil;

import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.ByteBuffer;
import java.nio.channels.FileChannel;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardOpenOption;
import java.util.Arrays;

/**
 * A workbook collects the strings of all text cells in a global list,
 * the Shared String Table. This table is located in the record SST in
 * the Workbook Globals Substream.
 * <p>
 * SST saves characters and strings sequentially. When writing a string,
 * it first determines whether it exists. If it exists, returns the index
 * in the Table (zero base), otherwise add it in to the last element of
 * Table and returns the current subscript.
 * Introduced Google BloomFilter to increase filtering speed, the
 * BloomFilter estimates the amount of data to be 1 million, and the false
 * positive rate is {@code 0.03%}. When the number exceeds {@code 2^20},
 * it will be converted to inline string.
 * <p>
 * A hot zone is also designed internally to cache multiple occurrences,
 * the default size is {@code 65,536}, and the LRU elimination algorithm is used.
 * If the cache misses, it will be read from in temp file and flushed to the
 * cache.
 * <p>
 * Characters are handled differently. ASCII characters use the built-in array
 * cache subscript. The over {@code 0x7F} characters will be converted to strings and
 * searched using strings.
 *
 * @author guanquan.wang on 2017/10/10.
 */
@TopNS(prefix = "", value = "sst", uri = Const.SCHEMA_MAIN)
public class SharedStrings implements Storageable, AutoCloseable {
    private Logger LOGGER = LogManager.getLogger(getClass());

    /**
     * The total word in workbook.
     */
    private int count;

    /**
     * Import google BloomFilter to check keyword not exists
     */
    private BloomFilter<String> filter;

    /**
     * Cache ASCII value
     */
    private int[] ascii;
    private Path temp;
    private ExtBufferedWriter writer;

    /**
     * Cache the string which read twice and above
     * Use LRU elimination algorithm
     */
    private Cache<String, Integer> hot;

    /**
     * Storage into temp file on disk
     */
    private SharedStringTable sst;

    private int j;
    // For debug
    private int total_char_cache, total_sst_find, total_hot;

    /**
     * The number of expected insertions to the constructed bloom
     */
    private int expectedInsertions = 1 << 20;

    SharedStrings() {
        hot = FixSizeLRUCache.create();
        ascii = new int[1 << 7];
        // -1 means the keyword not exists
        Arrays.fill(ascii, -1);
        // Create a 2^20 expected insertions and 0.03% fpp bloom filter
        filter = BloomFilter.create(Funnels.stringFunnel(StandardCharsets.UTF_8), expectedInsertions, 0.0003);

        init();
    }

    /**
     * Create a temp file to storage all text cells
     */
    private void init() {
        try {
            temp = Files.createTempFile("~", "sst");
            writer = new ExtBufferedWriter(Files.newBufferedWriter(temp, StandardCharsets.UTF_8));

            sst = new SharedStringTable();
        } catch (IOException e) {
            throw new ExcelWriteException(e);
        }
    }

    private ThreadLocal<char[]> charCache = ThreadLocal.withInitial(() -> new char[1]);

    /**
     * Getting the character value index (zero base)
     *
     * @param c the character value
     * @return the index in ShareString
     * @throws IOException if I/O error occur
     */
    public int get(char c) throws IOException {
        // An ASCII keyword
        if (c < 128) {
            int n = ascii[c];
            // Not exists
            if (n == -1) {
                n = add(c);
                ascii[c] = n;
            }
            count++;
            total_char_cache++;
            return n;
        } else {
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
     * @throws IOException if I/O error occur
     */
    public int get(String key) throws IOException {
        count++;
        // The keyword not exists
        if (!filter.mightContain(key)) {
            // Reset the filter
            if (j >= expectedInsertions) {
                resetBloomFilter();
            }
            // Add to bloom if not full
            filter.put(key);
            j++;
            return add(key);
        }
        // Check the keyword exists in cache
        Integer n = hot.get(key);
        if (n == null) {
            if (sst.size() <= expectedInsertions) {
                // Find in temp file
                n = sst.find(key);
                total_sst_find++;
                // Append to last and cache it
                if (n < 0) {
                    n = add(key);
                }
                hot.put(key, n);
            } else {
                // Convert to inline string
                n = add(key);
            }
        } else {
            total_hot++;
        }
        return n;
    }

    private int add(String key) throws IOException {
        writer.write("<si><t>");
        writer.escapeWrite(key);
        writer.write("</t></si>");

        // Add to table
        return sst.push(key);
    }

    private int add(char c) throws IOException {
        writer.write("<si><t>");
        writer.escapeWrite(c);
        writer.write("</t></si>");

        // Add to table
        return sst.push(c);
    }

    @Override
    public void writeTo(Path root) throws IOException {
        // Close temp writer
        FileUtil.close(writer);

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

            if (sst.size() > 0) {
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
        try (FileChannel tempChannel = FileChannel.open(temp, StandardOpenOption.READ)) {
            tempChannel.transferTo(0, tempChannel.size(), channel);
        }
    }

    /**
     * Reset the bloom filter
     */
    private void resetBloomFilter() {
//        expectedInsertions <<= 1;
        filter = BloomFilter.create(Funnels.stringFunnel(StandardCharsets.UTF_8), expectedInsertions, 0.0003);
//        for (String key : hot) {
//            filter.put(key);
//        }
        for (Cache.Entry<String, Integer> e : hot) {
            filter.put(e.getKey());
        }
        j = hot.size();
    }

    @Override
    public void close() throws IOException {
        LOGGER.debug("total: {}, hot: {}, sst: {}, cache: {}"
            , count, total_hot, total_sst_find, total_char_cache);
        filter = null;
        hot.clear();
        hot = null;
        sst.close();
        FileUtil.rm(temp);
    }
}

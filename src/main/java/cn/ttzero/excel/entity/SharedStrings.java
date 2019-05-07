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
 * 字符串共享，一个workbook的所有worksheet共享
 * <p>
 * Created by guanquan.wang on 2017/10/10.
 */
@TopNS(prefix = "", value = "sst", uri = Const.SCHEMA_MAIN)
public class SharedStrings {
    // 存储共享字符
    private int count; // workbook所有字符串(shared属性为true)的个数
    private int uniqueCount;
    // Import google BloomFilter to check keyword not exists
    private BloomFilter<String> filter;
    // Cache ASCII value
    private int[] ascii;
    private Path temp;
    private ExtBufferedWriter writer;
    private Hot<String, Integer> hot;

    private StringBuilder buf;

    SharedStrings() {
        hot = new Hot<>(1 << 10);
        ascii = new int[1 << 7];
        // -1 means the keyword not exists
        Arrays.fill(ascii, -1);
        // Create a 10w expected insertions and 0.03% fpp bloom filter
        filter = BloomFilter.create(Funnels.stringFunnel(StandardCharsets.UTF_8), 100_000, 0.0003);

        init();
    }

    private void init() {
        try {
            temp = Files.createTempFile("+", "sst");
            Files.deleteIfExists(temp);
            writer = new ExtBufferedWriter(Files.newBufferedWriter(temp, StandardCharsets.UTF_8));
        } catch (IOException e) {
            throw new ExcelWriteException(e);
        }
    }

    private ThreadLocal<char[]> charCache = ThreadLocal.withInitial(() -> new char[1]);

    /**
     * Getting the character value index
     * @param c the character value
     * @return the index in ShareString
     */
    public int get(char c) throws IOException {
        // An ASCII keyword
        if (c < 128) {
            int n = ascii[c];
            // Not exists
            if (n == -1) {
                add(c);
                n = uniqueCount++;
                ascii[c] = n;
            }
            count++;
            return n;
        } else {
            char[] cs = charCache.get();
            cs[0] = c;
            return get(new String(cs));
        }
    }

    /**
     * Getting the string value from cache
     * @param key the string value
     * @return index of the string in the SST
     */
    public int get(String key) throws IOException {
        count++;
        // The keyword not exists
        if (!filter.mightContain(key)) {
            filter.put(key);
            int n = uniqueCount++;
            add(key);
            return n;
        }
        // Check the keyword exists in cache
        Integer n = hot.get(key);
        if (n == null) {
            // Find in temp file
            n = findFromFile(key);
            if (n >= 0) {
                hot.push(key, n);
            } else add(key);
        }
        return n;
    }

    private void add(String key) throws IOException {
        writer.write("<si><t>");
        writer.escapeWrite(key);
        writer.write("</t></si>");
    }

    private void add(char c) throws IOException {
        writer.write("<si><t>");
        writer.write(c);
        writer.write("</t></si>");
    }

    public void write(Path root) throws IOException {
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
                .append(" uniqueCount=\"").append(uniqueCount).append("\">")
                .append(Const.lineSeparator);
        } else {
            buf.append("<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"")
                .append(count).append("\" uniqueCount=\"").append(uniqueCount).append("\">")
            .append(Const.lineSeparator);
        }

        // The output path
        Path dist = root.resolve(StringUtil.lowFirstKey(getClass().getSimpleName() + Const.Suffix.XML));

        try (FileOutputStream fos = new FileOutputStream(dist.toFile());
            FileChannel channel = fos.getChannel()) {
            ByteBuffer buffer = ByteBuffer.allocate(512);
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
        // destroy
        destroy();
    }

    /**
     * Transfer temp data to dist path
     * @param channel the dist file channel
     * @throws IOException if io error occur
     */
    private void transfer(FileChannel channel) throws IOException {
        try (FileChannel tempChannel = FileChannel.open(temp, StandardOpenOption.READ)) {
            tempChannel.transferTo(0, tempChannel.size(), channel);
        }
    }

    /**
     *
     * @param key the string value
     * @return the index in sst file (zero base)
     */
    private int findFromFile(String key) throws IOException {
        writer.flush();

        buf = new StringBuilder();

        try (BufferedReader reader = Files.newBufferedReader(temp, StandardCharsets.UTF_8)) {
            char[] cb = new char[8192];
            int n, off = 0;
            O o = new O();
            o.cb = cb;
            o.key = key;
            while ((n = reader.read(cb, off, cb.length - off)) > 0) {
                o.n = n;
                findT(o);
                // Get it
                if (o.f) {
                    return o.index;
                } else {
                    System.arraycopy(cb, o.off, cb, 0, off = n - o.off);
                }
            }
        }

        return -1;
    }

    private void findT(O o) {
        int sn = 7, ln = 9 + sn, nChar = sn;
        for (; nChar < o.n; ) {
            o.index++;
            int next = next(o.cb, nChar, o.n);
            if (next >= o.n) {
                o.off = nChar;
                break;
            }
            String v = unescape(buf, o.cb, nChar, next);
            o.f = v.equals(o.key);
            if (o.f) {
                break;
            }
            if (next + ln > o.n) {
                o.off = nChar;
                break;
            }
            nChar = next + ln;
        }
    }

    private int next(char[] cb, int nChar, int n) {
        for (; nChar < n && cb[nChar] != '<'; nChar++);
        return nChar;
    }

    /**
     * clear memory
     */
    private void destroy() {
        FileUtil.rm(temp);
        buf = null;
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

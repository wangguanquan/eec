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

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.UncheckedIOException;
import java.nio.ByteBuffer;
import java.nio.ByteOrder;
import java.nio.channels.FileChannel;
import java.nio.channels.SeekableByteChannel;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardOpenOption;
import java.util.Arrays;
import java.util.Iterator;

import static java.nio.charset.StandardCharsets.UTF_8;

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
 * positive rate is 0.03%. When the number exceeds 1 million, it will be
 * carried out, redistributing 2 times the length of the space, the maximum
 * length 2^26, When the number exceeds 2^26, it will be converted to inline
 * string.
 *
 * A hot zone is also designed internally to cache multiple occurrences,
 * the default size is 1024, and the LRU elimination algorithm is used.
 * If the cache misses, it will be read from in temp file and flushed to the
 * cache.
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
    private Hot<String, Integer> hot;

    /**
     * Storage into temp file on disk
     */
    private SharedStringTable sst;

    /**
     * The number of expected insertions to the constructed bloom
     */
    private int expectedInsertions = 1 << 20;

    /**
     * The insertion limit
     */
    private static final int limitInsertions = 1 << 26;

    SharedStrings() {
        hot = new Hot<>(1 << 10);
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
     */
    public int get(String key) throws IOException {
        count++;
        // The keyword not exists
        if (!filter.mightContain(key)) {
            // Resize if not out of insertion limit
            if (sst.size() >= expectedInsertions) {
                if (expectedInsertions < limitInsertions) {
                    resizeBloomFilter();
                } else return -1;
            }
            // Add to bloom if not full
            filter.put(key);
            return add(key);
        }
        // Check the keyword exists in cache
        Integer n = hot.get(key);
        if (n == null) {
            // Find in temp file
            n = sst.find(key);
            // Append to last and cache it
            if (n < 0) {
                n = add(key);
            }
            hot.push(key, n);
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
     * Resize the bloom filter
     */
    private void resizeBloomFilter() {
        expectedInsertions <<= 1;
        System.out.println("resize to " + expectedInsertions);
        filter = BloomFilter.create(Funnels.stringFunnel(StandardCharsets.UTF_8), expectedInsertions, 0.0003);
        for (String key : sst) {
            filter.put(key);
        }
    }

    @Override
    public void close() throws IOException {
        filter = null;
        hot.clear();
        hot = null;
        sst.close();
        FileUtil.rm(temp);
    }

    /**
     *
     */
    public static class SharedStringTable implements AutoCloseable, Iterable<String> {
        /**
         * The temp path
         */
        private Path temp;

        /**
         * The total unique word in workbook.
         */
        private int count;

        private SeekableByteChannel channel;

        /**
         * Byte array buffer
         */
        private ByteBuffer buffer;

        protected SharedStringTable() throws IOException {
            temp = Files.createTempFile("+", ".sst");
            channel = Files.newByteChannel(temp, StandardOpenOption.WRITE, StandardOpenOption.READ);
            buffer = ByteBuffer.allocate(1 << 11);
            buffer.order(ByteOrder.LITTLE_ENDIAN);
        }

        /**
         * Write character value into table
         *
         * @param c the character value
         * @return the value index of table
         * @throws IOException if io error occur
         */
        public int push(char c) throws IOException {
            if (buffer.remaining() < 8) {
                flush();
            }
            buffer.putInt(2);
            buffer.putShort((short) 0x8000);
            buffer.putChar(c);
            return count++;
        }

        /**
         * Write string value into table
         *
         * @param key the string value
         * @return the value index of table
         * @throws IOException if io error occur
         */
        public int push(String key) throws IOException {
            byte[] bytes = key.getBytes(UTF_8);
            if (buffer.remaining() < bytes.length + 6) {
                flush();
            }
            buffer.putInt(bytes.length);
            buffer.putShort((short) key.length());
            buffer.put(bytes);
            return count++;
        }

        /**
         * Find character value from begging
         *
         * @param c the character to find
         * @return the index of character in shared string table
         * @throws IOException if io error occur
         */
        public int find(char c) throws IOException {
            return find(c, 0);
        }

        /**
         * Find from the specified location
         *
         * @param c the character to find
         * @param pos the buffer's position
         * @return the index of character in shared string table
         * @throws IOException if io error occur
         */
        public int find(char c, int pos) throws IOException {
            // Flush before read
            flush();
            int index = 0;
            long position = channel.position();
            // Read at start position
            channel.position(pos);
            A: for (; ;) {
                int dist = channel.read(buffer);
                // EOF
                if (dist <= 0) break;
                buffer.flip();
                for (; buffer.remaining() >= 8 && hasFullValue(buffer);) {
                    int a = buffer.getInt();
                    short n = buffer.getShort();
                    // A char value
                    if (n == (short) 0x8000) {
                        // Get it
                        if (buffer.getChar() == c) {
                            break A;
                        }
                    } else buffer.position(buffer.position() + a);
                    index++;
                }
                buffer.compact();
            }
            channel.position(position);
            buffer.rewind();
            // Returns -1 if not found
            return index < count ? index : -1;
        }

        /**
         * Find value from begging
         *
         * @param key the key to find
         * @return the index of character in shared string table
         * @throws IOException if io error occur
         */
        public int find(String key) throws IOException {
            return find(key, 0);
        }

        /**
         * Find from the specified location
         *
         * @param key the key to find
         * @param pos the buffer's position
         * @return the index of character in shared string table
         * @throws IOException if io error occur
         */
        public int find(String key, int pos) throws IOException {
            // Flush before read
            flush();
            int index = 0;
            long position = channel.position();
            // Read at start position
            channel.position(pos);
            byte[] bytes = key.getBytes(UTF_8);
            A: for (; ;) {
                int dist = channel.read(buffer);
                // EOF
                if (dist <= 0) break;
                buffer.flip();
                for (; buffer.remaining() >= 8 && hasFullValue(buffer);) {
                    int a = buffer.getInt();
                    short n = buffer.getShort();
                    // A string value
                    if (n != (short) 0x8000 && n == key.length()) {
                        int i = 0;
                        for (; i < a; ) {
                            if (buffer.get() != bytes[i++]) break;
                        }
                        if (i < a) {
                            buffer.position(buffer.position() + a - i);
                        } else break A;
                    } else buffer.position(buffer.position() + a);
                    index++;
                }
                buffer.compact();
            }
            channel.position(position);
            buffer.rewind();
            // Returns -1 if not found
            return index < count ? index : -1;
        }

        /**
         * Returns the cache size
         *
         * @return total keyword
         */
        public int size() {
            return count;
        }

        /**
         * Write buffered data to channel
         *
         * @throws IOException if io error occur
         */
        private void flush() throws IOException {
            buffer.flip();
            channel.write(buffer);
            buffer.compact();
        }

        /**
         * Check the remaining data is complete
         *
         * @param buffer the ByteBuffer
         * @return true or false
         */
        protected static boolean hasFullValue(ByteBuffer buffer) {
            int position = buffer.position();
            int n = buffer.get(position)   & 0xFF;
            n |= (buffer.get(position + 1) & 0xFF) <<  8;
            n |= (buffer.get(position + 2) & 0xFF) << 16;
            n |= (buffer.get(position + 3) & 0xFF) << 24;
            return n + 6 <= buffer.remaining();
        }

        /**
         * Close channel and delete temp files
         *
         * @throws IOException if io error occur
         */
        @Override
        public void close() throws IOException {
            buffer = null;
            if (channel != null) {
                channel.close();
            }
            FileUtil.rm(temp);
        }

        /**
         * Returns this buffer's position.
         *
         * @return  The position of this buffer
         */
        protected int position() {
            return buffer.position();
        }

        /**
         * Returns an iterator over elements of type String
         *
         * @return an Iterator.
         */
        @Override
        public Iterator<String> iterator() {
            try {
                flush();
            } catch (IOException e) {
                throw new UncheckedIOException(e);
            }
            return new SSTIterator(temp);
        }

        private static class SSTIterator implements Iterator<String> {
            private SeekableByteChannel channel;
            private ByteBuffer buffer;
            private byte[] bytes;
            private SSTIterator(Path temp) {
                try {
                    channel = Files.newByteChannel(temp, StandardOpenOption.WRITE, StandardOpenOption.READ);
                    buffer = ByteBuffer.allocate(1 << 11);
                    buffer.order(ByteOrder.LITTLE_ENDIAN);
                    // Read ahead
                    channel.read(buffer);
                    buffer.flip();
                } catch (IOException e) {
                    throw new UncheckedIOException(e);
                }
                bytes = new byte[100];
            }
            @Override
            public boolean hasNext() {
                try {
                    if (buffer.remaining() < 6 || !hasFullValue(buffer)) {
                        buffer.compact();
                        channel.read(buffer);
                        buffer.flip();
                    }
                    return buffer.hasRemaining();
                } catch (IOException e) {
                    throw new UncheckedIOException(e);
                }
            }

            @Override
            public String next() {
                int a = buffer.getInt();
                if (a > bytes.length) {
                    bytes = new byte[a];
                }
                // skip
                buffer.position(buffer.position() + 2);
                buffer.get(bytes, 0, a);
                return new String(bytes, 0, a, UTF_8);
            }
        }

    }
}

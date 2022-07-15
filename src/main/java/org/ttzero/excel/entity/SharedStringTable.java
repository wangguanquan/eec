/*
 * Copyright (c) 2017-2019, guanquan.wang@yandex.com All Rights Reserved.
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

import org.ttzero.excel.util.FileUtil;

import java.io.Closeable;
import java.io.IOException;
import java.io.UncheckedIOException;
import java.nio.ByteBuffer;
import java.nio.ByteOrder;
import java.nio.InvalidMarkException;
import java.nio.channels.SeekableByteChannel;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardOpenOption;
import java.util.Iterator;

import static java.nio.charset.StandardCharsets.UTF_8;
import static org.ttzero.excel.reader.SharedStrings.tableSizeFor;
import static org.ttzero.excel.util.FileUtil.exists;

/**
 * @author guanquan.wang at 2019-05-10 20:04
 */
public class SharedStringTable implements Closeable, Iterable<String> {
    /**
     * The temp path
     */
    private final Path temp;

    /**
     * The total unique word in workbook.
     */
    private int count;

    private final SeekableByteChannel channel;

    /**
     * Byte array buffer
     */
    private ByteBuffer buffer;

    /**
     * The channel mark not buffer
     */
    private long mark = -1;

    /**
     * Delete the temp file if @author {@link SharedStringTable}
     */
    protected boolean shouldDelete;

    /**
     * Buffer size(4K)
     */
    protected int defaultBufferSize = 1 << 12;

    /**
     * Create a temp file to storage shared strings
     *
     * @throws IOException if I/O error occur.
     */
    protected SharedStringTable() throws IOException {
        temp = Files.createTempFile("+", ".sst");
        shouldDelete = true;
        channel = Files.newByteChannel(temp, StandardOpenOption.WRITE, StandardOpenOption.READ);
        buffer = ByteBuffer.allocate(defaultBufferSize);
        buffer.order(ByteOrder.LITTLE_ENDIAN);
        // Total keyword storage the header 4 bytes
        buffer.putInt(0);
        flush();
    }

    /**
     * Constructor a SharedStringTable with a exists file
     *
     * @param path the file path
     * @throws IOException if file not exists or I/O error occur.
     */
    protected SharedStringTable(Path path) throws IOException {
        if (!exists(path)) {
            throw new IOException("The index path [" + path + "] not exists.");
        }
        this.temp = path;
        channel = Files.newByteChannel(temp, StandardOpenOption.WRITE, StandardOpenOption.READ);

        buffer = ByteBuffer.allocate(defaultBufferSize);
        buffer.order(ByteOrder.LITTLE_ENDIAN);

        channel.read(buffer);
        buffer.flip();

        if (buffer.remaining() > 4) {
            this.count = buffer.getInt();
        }
        buffer.clear();
        // Mark EOF
        channel.position(channel.size());
    }

    protected Path getTemp() {
        return temp;
    }

    /**
     * Write character value into table
     *
     * @param c the character value
     * @return the value index of table
     * @throws IOException if io error occur
     */
    public int push(char c) throws IOException {
        return pushChar(c);
    }

    /**
     * Write string value into table
     *
     * @param key the string value
     * @return the value index of table
     * @throws IOException if io error occur
     */
    public int push(String key) throws IOException {
        int len;
        if (key == null || (len = key.length()) == 0) {
            return pushChar((char) 0xFFFF);
        }
        if (len == 1) {
            return pushChar(key.charAt(0));
        }
        byte[] bytes = key.getBytes(UTF_8);
        // The byte length exceeds 4k
        if (bytes.length + 4 > defaultBufferSize) {
           return -1;
        }
        if (buffer.remaining() < bytes.length + 4) {
            flush();
        }
        buffer.putInt(bytes.length);
        buffer.put(bytes);
        return count++;
    }

    /**
     * Write character value into table
     *
     * @param c the character value
     * @return the value index of table
     * @throws IOException if io error occur
     */
    private int pushChar(char c) throws IOException {
        if (buffer.remaining() < 4) {
            flush();
        }
        buffer.putInt(~c);
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
        return find(c, 0L);
    }

    /**
     * Find from the specified location
     *
     * @param c the character to find
     * @param pos the buffer's position
     * @return the index of character in shared string table
     * @throws IOException if io error occur
     */
    public int find(char c, long pos) throws IOException {
        // Flush before read
        flush();
        int index = 0;
        // Mark current position
        mark().skip(pos);

        A: for (; ;) {
            int dist = channel.read(buffer);
            // EOF
            if (dist <= 0) break;
            buffer.flip();
            for (; checkCapacityAndGrow(); ) {
                int a = buffer.getInt();
                // A char value
                if (a < 0) {
                    // Get it
                    if (~a == c) {
                        break A;
                    }
                } else buffer.position(buffer.position() + a);
                index++;
            }
            buffer.compact();
        }
        reset();
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
        return find(key, 0L);
    }

    /**
     * Find from the specified location
     *
     * @param key the key to find
     * @param pos the buffer's position
     * @return the index of character in shared string table
     * @throws IOException if io error occur
     */
    public int find(String key, long pos) throws IOException {
        // Flush before read
        flush();
        // Mark current position
        mark().skip(pos);

        int index;
        if (key != null && !key.isEmpty()) {
            index = findKey(key);
        } else {
            index = findNull();
        }

        reset();
        buffer.rewind();
        // Returns -1 if not found
        return index < count ? index : -1;
    }

    /**
     * Find null or empty value in SharedStringTable
     *
     * @return the index of null or empty value in the SharedStringTable
     * @throws IOException if I/O error occur.
     */
    private int findNull() throws IOException {
        int index = 0;
        int n = ~(char) 0xFFFF;
        A: for (; ;) {
            int dist = channel.read(buffer);
            // EOF
            if (dist <= 0) break;
            buffer.flip();
            for (; checkCapacityAndGrow(); ) {
                int a = buffer.getInt();
                // Found the first Null or Empty value
                if (a == n) break A;
                if (a > 0) {
                    // A string value
                    buffer.position(buffer.position() + a);
                }
                index++;
            }
            buffer.compact();
        }
        return index;
    }

    /**
     * Find value in SharedStringTable
     *
     * @param key the string key
     * @return the index in the SharedStringTable
     * @throws IOException if I/O error occur
     */
    private int findKey(String key) throws IOException {
        int index = 0;
        byte[] bytes = key.getBytes(UTF_8);
        A: for (; ;) {
            int dist = channel.read(buffer);
            // EOF
            if (dist <= 0) break;
            buffer.flip();
            for (; checkCapacityAndGrow(); ) {
                int a = buffer.getInt();
                // Character value
                if (a < 0) {
                    index++;
                    continue;
                }
                // A string value
                if (a == bytes.length) {
                    int i = 0, p = buffer.position(), b = a - 1;
                    for (; i <= b; b--, i++) {
                        if (bytes[b] != buffer.get(p + b) || buffer.get() != bytes[i]) break;
                    }
                    if (i < b || i == b && (a & 1) == 1) {
                        buffer.position(p + a);
                    } else break A;
                } else buffer.position(buffer.position() + a);
                index++;
            }
            buffer.compact();
        }
        return index;
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
        if (buffer.hasRemaining()) {
            channel.write(buffer);
        }
        buffer.clear();
    }

    /**
     * Check remaining data and grow if shortage
     *
     * @return true/false
     */
    protected boolean checkCapacityAndGrow() {
        int i = hasFullValue(buffer);
        if (i < 0) {
            this.buffer = grow(buffer);
            this.buffer.flip();
        }

        return i > 0;
    }

    /**
     * Grow a new Buffer
     *
     * @param buffer Old buffer
     * @return new capacity buffer at {@code WRITE STATUS}
     */
    public static ByteBuffer grow(ByteBuffer buffer) {
        int n = nextByteSize(buffer) + 4;
        int newCapacity = Math.max(tableSizeFor(n), buffer.limit() << 1);
        ByteBuffer newBuffer = ByteBuffer.allocate(newCapacity);
        newBuffer.order(ByteOrder.LITTLE_ENDIAN);
        newBuffer.put(buffer);
        return newBuffer;
    }

    /**
     * Check the remaining data is complete
     *
     * @param buffer the ByteBuffer
     * @return 1: full 0: shortage -1: The length of the word exceeds the Buffer
     */
    public static int hasFullValue(ByteBuffer buffer) {
        int n = nextByteSize(buffer);
        return n < 0 || n + 4 <= buffer.remaining() ? 1 : n > buffer.limit() - 4 ? -1 : 0;
    }

    /**
     * Calculate the size of the next block of text
     * <p>
     * Note: The remaining bytes size must be judged before calling this method
     *
     * @param buffer Data Buffer
     * @return block size
     */
    static int nextByteSize(ByteBuffer buffer) {
        if (buffer.remaining() < 4) return 0;
        int position = buffer.position();
        int n = buffer.get(position)   & 0xFF;
        n |= (buffer.get(position + 1) & 0xFF) <<  8;
        n |= (buffer.get(position + 2) & 0xFF) << 16;
        n |= (buffer.get(position + 3) & 0xFF) << 24;
        return n;
    }

    /**
     * Commit current index file writer
     *
     * @throws IOException if I/O error occur
     */
    protected void commit() throws IOException {
        flush();
        buffer.putInt(count);
        buffer.flip();
        channel.position(0);
        channel.write(buffer);
    }

    /**
     * Close channel and delete temp files
     *
     * @throws IOException if I/O error occur
     */
    @Override
    public void close() throws IOException {
        // Commit writer
        commit();
        // Release
        buffer = null;
        if (channel != null) {
            channel.close();
        }
        if (shouldDelete) {
            FileUtil.rm(temp);
        }
    }

    /**
     * Returns this buffer's position.
     *
     * @return  The position of this buffer
     * @throws IOException if I/O error occur
     */
    protected long position() throws IOException {
        return channel.position() + buffer.position();
    }

    /**
     * Returns a ByteBuffer data from channel position
     *
     * @param buffer the byte buffer
     * @return the read data length
     * @throws IOException if I/O error occur
     */
    protected int read(ByteBuffer buffer) throws IOException {
        return channel.read(buffer);
    }

    /**
     * Sets this buffer's mark at its position.
     *
     * @return  This SharedStringTable
     * @throws IOException if I/O error occur
     */
    protected SharedStringTable mark() throws IOException {
        flush();
        mark = channel.position();
        return this;
    }

    /**
     * Resets this buffer's position to the previously-marked position.
     *
     * Invoking this method neither changes nor discards the mark's
     * value.
     *
     * @return  This SharedStringTable
     *
     * @throws  InvalidMarkException If the mark has not been set
     * @throws IOException if I/O error occur
     */
    protected SharedStringTable reset() throws IOException {
        if (mark == -1)
            throw new InvalidMarkException();
        channel.position(mark);
        mark = -1;
        buffer.clear();
        return this;
    }

    /**
     * Jump to the specified position, the actual moving position
     * will be increased by 4, the header contains an integer value.
     *
     * @param position the position to jump
     * @return the {@link SharedStringTable}
     * @throws IOException if I/O error occur
     */
    protected SharedStringTable skip(long position) throws IOException {
        channel.position(position + 4);
        return this;
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
        private final SeekableByteChannel channel;
        private ByteBuffer buffer;
        private byte[] bytes;
        @SuppressWarnings("unused")
        private int count; // ignore
        private int fv; // Full Value Mark
        private final char[] chars;
        private SSTIterator(Path temp) {
            try {
                channel = Files.newByteChannel(temp, StandardOpenOption.READ);
                buffer = ByteBuffer.allocate(1 << 11);
                buffer.order(ByteOrder.LITTLE_ENDIAN);
                // Read ahead
                channel.read(buffer);
                buffer.flip();
                if (buffer.remaining() > 4) {
                    count = buffer.getInt();
                }
            } catch (IOException e) {
                throw new UncheckedIOException(e);
            }
            bytes = new byte[128];
            chars = new char[1];
        }
        @Override
        public boolean hasNext() {
            try {
                if (buffer.remaining() < 6 || (fv = hasFullValue(buffer)) <= 0) {
                    if (fv < 0) {
                        this.buffer = grow(buffer);
                    }
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
            if (a < 0) {
                char c = (char) ~a;
                if (c < 0xFFFF) {
                    chars[0] = c;
                    return new String(chars);
                } else return "";
            } else {
                buffer.get(bytes, 0, a);
                return new String(bytes, 0, a, UTF_8);
            }
        }
    }

}

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

import cn.ttzero.excel.util.FileUtil;

import java.io.IOException;
import java.io.UncheckedIOException;
import java.nio.ByteBuffer;
import java.nio.ByteOrder;
import java.nio.channels.SeekableByteChannel;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardOpenOption;
import java.util.Iterator;

import static java.nio.charset.StandardCharsets.UTF_8;

/**
 * Create by guanquan.wang at 2019-05-08 15:31
 */
class SharedStringTable implements AutoCloseable, Iterable<String> {
    /**
     * The temp path
     */
    private Path temp;

    /**
     * Searches for only the first 100,000 words
     */
    private int find_limit = 100_000;

    /**
     * The total unique word in workbook.
     */
    private int count;

    private SeekableByteChannel channel;

    /**
     * Byte array buffer
     */
    private ByteBuffer buffer;

    SharedStringTable() {
        try {
            temp = Files.createTempFile("+", ".sst");
            channel = Files.newByteChannel(temp, StandardOpenOption.WRITE, StandardOpenOption.READ);
            buffer = ByteBuffer.allocate(1 << 11);
            buffer.order(ByteOrder.LITTLE_ENDIAN);
        } catch (IOException e) {
            throw new ExcelWriteException(e);
        }
    }

    public int push(char c) throws IOException {
        if (buffer.remaining() < 8) {
            flush();
        }
        buffer.putInt(2);
        buffer.putShort((short) 0x8000);
        buffer.putChar(c);
        return count++;
    }

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

    public int find(char c) throws IOException {
        // Flush before read
        flush();
        int index = 0;
        long position = channel.position();
        // Read at start position
        channel.position(0);
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

    public int find(String key) throws IOException {
        // Flush before read
        flush();
        int index = 0;
        long position = channel.position();
        // Read at start position
        channel.position(0);
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
     * @return total keyword
     */
    public int size() {
        return count;
    }

    /**
     * Write buffered data to channel
     * @throws IOException if io error occur
     */
    private void flush() throws IOException {
        buffer.flip();
        channel.write(buffer);
        buffer.compact();
    }

    private /* FOR TESTS */ void print(byte[] bytes, int off, int len) {
        for (int i = off, n = 0; i < len; i++) {
            System.out.print(bytes[i] & 0xFF);
            System.out.print(' ');
            if (++n % 64 == 0) System.out.println();
        }
    }

    /**
     * Check the remaining data is complete
     * @param buffer the ByteBuffer
     * @return true or false
     */
    private static boolean hasFullValue(ByteBuffer buffer) {
        int position = buffer.position();
        int n = buffer.get(position)   & 0xFF;
        n |= (buffer.get(position + 1) & 0xFF) <<  8;
        n |= (buffer.get(position + 2) & 0xFF) << 16;
        n |= (buffer.get(position + 3) & 0xFF) << 24;
        return n + 6 <= buffer.remaining();
    }

    /**
     * Close channel and delete temp files
     * @throws IOException if io error occur
     */
    @Override
    public void close() throws IOException {
        if (channel != null) {
            channel.close();
        }
        FileUtil.rm(temp);
    }

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
            buffer.getShort();
            if (a > bytes.length) {
                bytes = new byte[a];
            }
            buffer.get(bytes, 0, a);
            return new String(bytes, 0, a, UTF_8);
        }
    }

}

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
import cn.ttzero.excel.entity.SharedStringTable;
import cn.ttzero.excel.util.FileUtil;

import java.io.IOException;
import java.nio.ByteBuffer;
import java.nio.ByteOrder;
import java.nio.channels.SeekableByteChannel;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardOpenOption;

import static java.nio.charset.StandardCharsets.UTF_8;

/**
 * Create by guanquan.wang at 2019-05-10 20:06
 */
public class IndexSharedStringTable extends SharedStringTable {
    /**
     * The index temp path
     */
    private Path temp;

    private SeekableByteChannel channel;

    /**
     * Byte array buffer
     */
    private ByteBuffer buffer, readBuffer;

    /**
     * The sector size
     */
    private int sst = 8;

    /**
     * The short sector size
     */
    private int ssst = 6;

    /**
     * Record position once every 64 strings
     */
    private int kSplit = 0x7FFFFFFF >> ssst << ssst;

    /**
     * Flush buffer each 256 keys
     */
    private int kFlush = 0x7FFFFFFF >> sst << sst;

    /**
     * A multiplexing byte array
     */
    private byte[] bytes;

    /**
     * A fix length multiplexing char array
     */
    private char[] chars = new char[1];

    /**
     * Cache the getting index
     */
    private int index = -1;

    /**
     * Current read/write status
     */
    private byte status;

    private static final byte READ = 1, WRITE = 0;

    IndexSharedStringTable() throws IOException {
        super();

        Path superPath = getTemp();
        temp = Files.createFile(Paths.get(superPath.toString() + ".idx"));
        channel = Files.newByteChannel(temp, StandardOpenOption.WRITE, StandardOpenOption.READ);
        buffer = ByteBuffer.allocate(1 << 11);
        buffer.order(ByteOrder.LITTLE_ENDIAN);
        readBuffer = ByteBuffer.allocate(1 << 10);
        readBuffer.order(ByteOrder.LITTLE_ENDIAN);
    }

    IndexSharedStringTable(Path path) throws IOException {
        super(Paths.get(path.toString().substring(0, path.toString().length() - 4)));
        temp = path;
        channel = Files.newByteChannel(temp, StandardOpenOption.WRITE, StandardOpenOption.READ);
        channel.position(channel.size());
        buffer = ByteBuffer.allocate(1 << 11);
        buffer.order(ByteOrder.LITTLE_ENDIAN);
        readBuffer = ByteBuffer.allocate(1 << 10);
        readBuffer.order(ByteOrder.LITTLE_ENDIAN);
    }

    /**
     * Write character value into table
     *
     * @param c the character value
     * @return the value index of table
     * @throws IOException if io error occur
     */
    @Override
    public int push(char c) throws IOException {
        putsIndex();
        return super.push(c);
    }

    /**
     * Write string value into table
     *
     * @param key the string value
     * @return the value index of table
     * @throws IOException if io error occur
     */
    @Override
    public int push(String key) throws IOException {
        putsIndex();
        return super.push(key);
    }

    /**
     * Getting by index
     *
     * @param index the value's index in table
     * @return the string value at index
     */
    public String get(int index) throws IOException {
        checkBound(index);
        if (status == WRITE || index != this.index) {
            status = READ;
            long position = getIndexPosition(index);
            readBuffer.clear();

            super.mark();

            super.skip(position);

            int dist = read(readBuffer);

//            super.reset();

            if (dist < 0) {
                // TODO reader more data into index file
                return null;
            }
            readBuffer.flip();

            skipTo(index);
        } else if (!hasFullValue(readBuffer)) {
            readBuffer.compact();
            int dist = read(readBuffer);
            if (dist < 0) {
                return null;
            }
            readBuffer.flip();
        }

        if (hasFullValue(readBuffer)) {
            this.index = index + 1;
            return parse(readBuffer);
        }
        return null;
    }

    /**
     * Batch getting
     *
     * @param fromIndex the index of the first element, inclusive, to be sorted
     * @param array Destination array
     * @return The number of string read, or -1 if the end of the
     *              stream has been reached
     */
    public int get(int fromIndex, String[] array) throws IOException {
        checkBound(fromIndex);
        if (status == WRITE || fromIndex != this.index) {
            status = READ;
            long position = getIndexPosition(fromIndex);
            readBuffer.clear();

            super.mark();

            super.skip(position);

            int dist = read(readBuffer);
            if (dist < 0) {
                return 0;
            }
            readBuffer.flip();
        }
        int i = 0;
        A: for (; ;) {

            if (i == 0 && fromIndex != this.index) {
                skipTo(fromIndex);
            }

            for ( ; hasFullValue(readBuffer); ) {
                array[i++] = parse(readBuffer);
                if (i >= array.length) break A;
            }

            readBuffer.compact();

            int dist = read(readBuffer);
            if (dist < 0) {
                break;
            }
            readBuffer.flip();
        }

        this.index = fromIndex + i;

        return i;
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
     * Puts the main's position into index file if need.
     */
    private void putsIndex() throws IOException {
        int size = size();
        if (size > 0 && (size & kSplit) == size) {
            if ((size & kFlush) == size) {
                flush();
            }
            buffer.putLong(super.position() - 4);
        }
        // Check status
        if (status == READ) {
            status = WRITE;
            super.reset();
        }
    }

    /**
     * Check the getting index
     *
     * @param index the getting index
     */
    private void checkBound(int index) {
        int size = size();
        if (size <= index) {
            // TODO reader more data into index file
            throw new ExcelWriteException("index: " + index + ", size: " + size);
        }
    }

    private long getIndexPosition(int keyIndex) throws IOException {
        long position = 0L;
        if (keyIndex < (1 << ssst)) return position;
        long index_size = channel.position();
        // Read from file
        if (index_size >> 3 > keyIndex) {
            flush();
            long pos = channel.position();
            channel.position(keyIndex >> ssst << 3);
            channel.read(readBuffer);
            readBuffer.flip();
            position = readBuffer.getLong();
            channel.position(pos);

            // Read from buffer
        } else {
            buffer.flip();
            if (buffer.hasRemaining()) {
                buffer.mark();
                buffer.position((int) (keyIndex - index_size) >> ssst << 3);
                position = buffer.getLong();
                buffer.reset();
            }
            buffer.limit(buffer.capacity());
        }

        return position;
    }

    private String parse(ByteBuffer readBuffer) {
        int n = readBuffer.getInt();
        if (bytes == null || bytes.length < n) {
            bytes = new byte[n < 128 ? 128 : n];
        }
        if (n < 0) {
            chars[0] = (char) ~n;
            return new String(chars);
        } else {
            readBuffer.get(bytes, 0, n);
            return new String(bytes, 0, n, UTF_8);
        }
    }

    private void skipTo(int index) {
        for (int n, i = index >> ssst << ssst; i < index; i++) {
            n = readBuffer.getInt();
            readBuffer.position(readBuffer.position() + 2 + n);
        }
    }

    @Override
    public void close() throws IOException {
        buffer = null;
        readBuffer = null;
        if (channel != null) {
            channel.close();
        }
        if (shouldDelete) {
            FileUtil.rm(temp);
        }

        super.close();
    }
}

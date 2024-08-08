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

package org.ttzero.excel.util;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.Closeable;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.OutputStreamWriter;
import java.io.UncheckedIOException;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.NoSuchElementException;
import java.util.Spliterator;
import java.util.Spliterators;
import java.util.function.Consumer;
import java.util.stream.Stream;
import java.util.stream.StreamSupport;

import static org.ttzero.excel.util.ExtBufferedWriter.MIN_INTEGER_CHARS;
import static org.ttzero.excel.util.ExtBufferedWriter.MIN_LONG_CHARS;
import static org.ttzero.excel.util.ExtBufferedWriter.getChars;
import static org.ttzero.excel.util.ExtBufferedWriter.stringSize;
import static org.ttzero.excel.util.FileUtil.exists;
import static org.ttzero.excel.util.FileUtil.mkdir;

/**
 * Comma-Separated Values
 * <p>
 * RFC 4180 standard
 * Reliance on the standard documented by RFC 4180 can simplify CSV exchange. However,
 * this standard only specifies handling of text-based fields.
 * Interpretation of the text of each field is still application-specific.
 * <p>
 * RFC 4180 formalized CSV. It defines the MIME type "text/csv", and CSV files that follow its
 * rules should be very widely portable. Among its requirements:
 * <ul>
 * <li>MS-DOS-style lines that end with (CR/LF) characters (optional for the last line).</li>
 * <li>An optional header record (there is no sure way to detect whether it is present,
 * so care is required when importing).</li>
 * <li>Each record "should" contain the same number of comma-separated fields.</li>
 * <li>Any field may be quoted (with double quotes).</li>
 * <li>Fields containing a line-break, double-quote or commas should be quoted. (If they are not,
 * the file will likely be impossible to process correctly).</li>
 * <li>A (double) quote character in a field must be represented by two (double) quote characters.</li>
 * </ul>
 * @author guanquan.wang at 2019-02-12 17:27
 */
public class CSVUtil {
    private static final Logger LOGGER = LoggerFactory.getLogger(CSVUtil.class);

    private CSVUtil() { }
    private static final char QUOTE = '"';
    private static final char HT = 9;
    private static final char LF = 10;
    private static final char CR = 13;
    private static final char COMMA = ',';
    private static final String EMPTY = "";

    // --- Read

    /**
     * Read csv format file by UTF-8 charset.
     *
     * @param path the csv file path
     * @param clazz the class convert to
     * @param <T> the result type
     * @return ArrayList of clazz
     * @throws IOException if I/O error occur
     */
    public static <T> List<T> read(Path path, Class<T> clazz) throws IOException {
        return read(path, clazz, null);
    }

    /**
     * Read csv format file by UTF-8 charset.
     *
     * @param path the csv file path
     * @return ArrayList of string
     * @throws IOException if I/O error occur
     */
    public static List<String[]> read(Path path) throws IOException {
        return read(path, (Charset) null);
    }

    /**
     * Read csv format file by UTF-8 charset.
     *
     * @param path the csv file path
     * @param separator the separator character
     * @return ArrayList of string
     * @throws IOException if I/O error occur
     */
    public static List<String[]> read(Path path, char separator) throws IOException {
        return read(path, separator, null);
    }

    /**
     * Read csv format file.
     *
     * @param path the csv file path
     * @param clazz the class convert to
     * @param charset the charset to use for encoding
     * @param <T> the result type
     * @return ArrayList of clazz
     * @throws IOException if I/O error occur
     */
    public static <T> List<T> read(Path path, Class<T> clazz, Charset charset) throws IOException {
        throw new UnsupportedOperationException();
    }

    /**
     * Read csv format file.
     *
     * @param path the csv file path
     * @param charset the charset to use for encoding
     * @return ArrayList of string
     * @throws IOException if I/O error occur
     */
    public static List<String[]> read(Path path, Charset charset) throws IOException {
        return read(path, (char) 0x0, charset);
    }

    /**
     * Read csv format file.
     *
     * @param path the csv file path
     * @param separator the separator character
     * @param charset the charset to use for encoding
     * @return ArrayList of string
     * @throws IOException if I/O error occur
     */
    public static List<String[]> read(Path path, char separator, Charset charset) throws IOException {
        // Check comma character and column
        // FileNotFoundException will be occurred
        O o = init(path, separator, charset);
        // Empty file
        if (o == null) {
            return Collections.emptyList();
        }

        // Use iterator
        try (RowsIterator iter = new RowsIterator(o, path, o.charset)) {
            List<String[]> result = new ArrayList<>();
            while (iter.hasNext()) {
                result.add(iter.next());
            }
            return result;
        }
    }

    /**
     * Create a CSV reader.
     *
     * @param path the csv file path
     * @return a stream CSV format reader
     */
    public static Reader newReader(Path path) {
        return newReader(path, null);
    }

    /**
     * Create a CSV reader.
     *
     * @param path the csv file path
     * @param charset the charset to use for encoding
     * @return a stream CSV format reader
     */
    public static Reader newReader(Path path, Charset charset) {
        return new Reader(path, charset);
    }

    /**
     * Create a CSV reader.
     *
     * @param path the csv file path
     * @param separator the separator character
     * @return a stream CSV format reader
     */
    public static Reader newReader(Path path, char separator) {
        Reader reader = newReader(path);
        reader.separator = separator;
        return reader;
    }

    /**
     * Create a CSV reader.
     *
     * @param path the csv file path
     * @param separator the separator character
     * @param charset the charset to use for encoding
     * @return a stream CSV format reader
     */
    public static Reader newReader(Path path, char separator, Charset charset) {
        Reader reader = newReader(path, charset);
        reader.separator = separator;
        return reader;
    }

    // --- Writer

    /**
     * Save vector object as csv format file
     *
     * @param data the vector object to be save
     * @param path the save path
     * @throws IOException if I/O error occur
     */
    public static void writeTo(List<?> data, Path path) throws IOException {
        throw new UnsupportedOperationException();
    }

    /**
     * Create a CSV writer
     *
     * @param path the storage path
     * @return a CSV format writer
     * @throws IOException no permission or other I/O error occur
     */
    public static Writer newWriter(Path path) throws IOException {
        testOrCreate(path);
        return new Writer(path);
    }

    /**
     * Create a CSV writer
     *
     * @param path the storage path
     * @param charset the charset to use for encoding
     * @return a CSV format writer
     * @throws IOException no permission or other I/O error occur
     */
    public static Writer newWriter(Path path, Charset charset) throws IOException {
        testOrCreate(path);
        return new Writer(path, charset);
    }

    /**
     * Create a CSV writer
     *
     * @param path the storage path
     * @param separator the separator character
     * @return a CSV format writer
     * @throws IOException no permission or other I/O error occur
     */
    public static Writer newWriter(Path path, char separator) throws IOException {
        testOrCreate(path);
        Writer writer = new Writer(path);
        writer.separator = separator;
        return writer;
    }

    /**
     * Create a CSV writer
     *
     * @param path the storage path
     * @param separator the separator character
     * @param charset the charset to use for encoding
     * @return a CSV format writer
     * @throws IOException no permission or other I/O error occur
     */
    public static Writer newWriter(Path path, char separator, Charset charset) throws IOException {
        testOrCreate(path);
        Writer writer = new Writer(path, charset);
        writer.separator = separator;
        return writer;
    }

    /**
     * Create a CSV writer
     *
     * @param writer the {@link BufferedWriter}
     * @return a CSV format writer
     */
    public static Writer newWriter(BufferedWriter writer) {
        return new Writer(writer);
    }

    /**
     * Create a CSV writer
     *
     * @param os the {@link OutputStream}
     * @return a CSV format writer
     */
    public static Writer newWriter(OutputStream os) {
        return new Writer(new BufferedWriter(new OutputStreamWriter(os, StandardCharsets.UTF_8)));
    }

    private static void testOrCreate(Path path) throws IOException {
        if (!exists(path)) {
            mkdir(path.getParent());
        }
    }

    // --PUBLIC inner reader

    /**
     * A CSV format file reader.
     */
    public static class Reader implements Closeable {

        private RowsIterator iterator;
        private final Path path;
        private final Charset charset;
        private char separator;

        private Reader(Path path, Charset charset) {
            this.path = path;
            this.charset = charset;
            this.separator = (char) 0x0;
        }

        /**
         * Read csv format file.
         *
         * @param clazz the class convert to
         * @param <T> the result type
         * @return a stream of clazz array
         */
        public <T> Stream<T> stream(Class<T> clazz) {
            throw new UnsupportedOperationException();
        }

        /**
         * Read csv format file.
         *
         * @return a stream of string array
         * @throws IOException file not exists or read file error.
         */
        public Stream<String[]> stream() throws IOException {
            // Check comma character and column
            // FileNotFoundException will be occurred
            O o = init(path, separator, charset);
            // Empty file
            if (o == null) {
                return StreamSupport.stream(emptySql, false);
            }

            // Use iterator
            iterator = new RowsIterator(o, path, o.charset);

            return StreamSupport.stream(Spliterators.spliteratorUnknownSize(
                iterator, Spliterator.ORDERED | Spliterator.NONNULL), false);
        }

        /**
         * Read csv format file.
         * there has only one string array in memory, so do not call 'collect' or 'toArray' function direct.
         *
         * @return a stream of sheared string array
         * @throws IOException file not exists or read file error.
         */
        public Stream<String[]> sharedStream() throws IOException {
            return sharedStream((char) 0x0);
        }

        /**
         * Read csv format file.
         * there has only one string array in memory, so do not call 'collect' or 'toArray' function direct.
         *
         * @param separator the separator character
         * @return a stream of sheared string array
         * @throws IOException file not exists or read file error.
         */
        public Stream<String[]> sharedStream(char separator) throws IOException {
            // Check comma character and column
            // FileNotFoundException will be occurred
            O o = init(path, separator, charset);
            // Empty file
            if (o == null) {
                return StreamSupport.stream(emptySql, false);
            }

            // Use iterator
            iterator = new SharedRowsIterator(o, path, charset);

            return StreamSupport.stream(Spliterators.spliteratorUnknownSize(
                iterator, Spliterator.ORDERED | Spliterator.NONNULL), false);
        }

        /**
         * Read csv format file.
         *
         * @return an iterator
         * @throws IOException file not exists or read file error.
         */
        public RowsIterator iterator() throws IOException {
            // Check comma character and column
            // FileNotFoundException will be occur
            O o = init(path, separator, charset);
            // Empty file
            if (o == null) {
                return RowsIterator.createEmptyIterator();
            }

            // Use iterator
            return new RowsIterator(o, path, o.charset);
        }

        /**
         * Read csv format file.
         *
         * @return an iterator
         * @throws IOException file not exists or read file error.
         */
        public RowsIterator sharedIterator() throws IOException {
            // Check comma character and column
            // FileNotFoundException will be occur
            O o = init(path, separator, charset);
            // Empty file
            if (o == null) {
                return SharedRowsIterator.createEmptyIterator();
            }

            // Use iterator
            return new SharedRowsIterator(o, path, charset);
        }

        @Override
        public void close() throws IOException {
            if (iterator != null) {
                try {
                    iterator.close();
                } catch (Exception e) {
                    throw new IOException(e);
                }
            }
        }

        // None item iterator
        private final Spliterator<String[]> emptySql = new Spliterator<String[]>() {
            @Override
            public boolean tryAdvance(Consumer<? super String[]> action) {
                return false;
            }

            @Override
            public Spliterator<String[]> trySplit() {
                return null;
            }

            @Override
            public long estimateSize() {
                return 0;
            }

            @Override
            public int characteristics() {
                return 0;
            }
        };
    }

    // -- PRIVATE inner function

    private static class O {
        int offset, line;
        String value;
        boolean newLine;
        Charset charset;
        O(int offset) { this.offset = offset; }
    }

    /**
     * Rows iterator
     */
    public static class RowsIterator implements Closeable, Iterator<String[]> {
        private int column;
        private final char comma;
        private BufferedReader reader;
        private char[] chars;
        private int offset;
        private int i, _i;
        private int n;
        String[] nextRow;
        private static final int length = 8192;
        private O o;
        boolean EOF, load;

        RowsIterator() {
            this.comma = COMMA;
        }

        RowsIterator(O o, Path path, Charset charset) throws IOException {
            this.column = o.offset;
            this.comma = o.value.charAt(0);
            this.o = o;
            // Default charset UTF-8
            reader = charset != null ? Files.newBufferedReader(path, charset) : Files.newBufferedReader(path);
            // Ignore the Byte-order mark (BOM)
            if (o.line > 0) {
                reader.skip(o.line);
                o.line = 0;
            }
            chars = new char[length];
            nextRow = new String[column];
            this.offset = o.offset = 0;
            load = true;
        }

        @Override
        public boolean hasNext() {
            if (EOF) return false;
            try {
                for ( ; ; ) {
                    if (load) {
                        n = reader.read(chars, offset, length - offset);
                        // EOF
                        if (n <= 0) {
                            EOF = true;
                            // End of {comma}
                            if (chars[o.offset - 1] == comma) {
                                // Contain more than standard comma-separated fields
                                if (i == column) {
                                    String[] _array = new String[++column];
                                    System.arraycopy(nextRow, 0, _array, 0, column - 1);
                                    nextRow = _array;
                                }
                                nextRow[i++] = EMPTY;
                                _i = i;
                            }
                            return nextRow[0] != null;
                        }
                        n += offset;
                        o.offset = 0;
                        load = false;
                    }
                    // Parse a block characters
                    while (parse(chars, n, o, comma)) {
                        offset = o.offset;
                        // Contain more than standard comma-separated fields
                        if (i == column) {
                            String[] _array = new String[++column];
                            System.arraycopy(nextRow, 0, _array, 0, column - 1);
                            nextRow = _array;
                        }
                        nextRow[i++] = o.value;
                        _i = i;
                        // End of block
                        if (offset >= n) {
                            // An integral row
                            if (o.newLine) {
                                o.line++;
                                // Ignore empty row
                                if ((i > 1 || nextRow[0] != null)) {
                                    // Line end of '{comma}'
                                    if (o.value == null) nextRow[i - 1] = EMPTY;
                                    i = 0;
                                    load = true;
                                    offset = 0;
                                    return load;
                                }
                            }
                            break;
                        }
                        if (!o.newLine) continue;
                        // Unquoted (CR/LF) characters
                        o.line++;
                        // Ignore empty row
                        if ((i > 1 || nextRow[0] != null)) {
                            // Line end of '{comma}'
                            if (o.value == null) nextRow[i - 1] = EMPTY;
                            i = 0;
                            return true;
                        }
                        i = 0;
                    }
                    load = true;
                    // Move the last character to header
                    if (offset < n) {
                        System.arraycopy(chars, offset, chars, 0, offset = n - offset);
                    } else offset = 0;
                }
            } catch (IOException e) {
                throw new UncheckedIOException(e);
            }
        }

        @Override
        public String[] next() {
            if (nextRow[0] != null || hasNext()) {
                String[] next = Arrays.copyOf(nextRow, _i);
                nextRow[0] = null;
                return next;
            } else {
                throw new NoSuchElementException();
            }
        }

        @Override
        public void close() throws IOException {
            if (reader != null) {
                reader.close();
            }
            chars = null;
            nextRow = null;
        }

        static RowsIterator createEmptyIterator() {
            RowsIterator iterator = new RowsIterator();
            iterator.EOF = true;
            iterator.nextRow = new String[0];
            return iterator;
        }
    }

    /**
     * Shared Row iterator
     */
    public static class SharedRowsIterator extends RowsIterator {
        /*
         A flag to mark the next row is ready.
         */
        private boolean produced;

        protected SharedRowsIterator() {
            super();
        }

        SharedRowsIterator(O o, Path path, Charset charset) throws IOException {
            super(o, path, charset);
        }

        @Override
        public boolean hasNext() {
            if (produced) return true;
            nextRow[0] = null;
            return produced = super.hasNext();
        }

        @Override
        public String[] next() {
            if (produced || hasNext()) {
                produced = false;
                return nextRow;
            } else {
                throw new NoSuchElementException();
            }
        }

        /**
         * Retain current row data
         */
        public void retain() {
            produced = true;
        }
    }

    /**
     * Check comma character and column
     *
     * @param path the csv file path
     * @param separator the separator character
     * @param charset the charset to use for encoding
     * @return comma character and column size
     */
    private static O init(Path path, char separator, Charset charset) throws IOException {
        // Test charset
        Charset bom = charsetTest(path);
        // Default use UTF-8 charset
        if (bom == null && charset == null) {
            charset = StandardCharsets.UTF_8;
        }
        if (bom != null) {
            if (charset == null) {
                charset = bom;
                // Print a warring log or reset charset
            } else if (!charset.equals(bom)) {
                LOGGER.warn("Maybe the charset is " + bom);
            }
        }

        try (BufferedReader reader = Files.newBufferedReader(path, charset)) {
            int n = 0; // read 10 lines
            String[] lines = new String[10];
            String s;
            while ((s = reader.readLine()) != null && n < lines.length) {
                if (!s.isEmpty()) {
                    lines[n++] = s;
                }
            }

            // Empty file
            if (lines[0] == null || lines[0].isEmpty()) {
                return null;
            }
            // No enough information to judge the separator
            if (n < 10) {
                LOGGER.warn("No enough information to judge the separator.");
            }

            // USA/UK CSV file almost use ',' or '\t'
            // European CSV file almost use ';'
            String[] commas = separator == 0
                ? new String[] { String.valueOf(COMMA), String.valueOf(HT), ";" } : new String[] { String.valueOf(separator)};
            int[][] columns = new int[commas.length][n];
            for (int i = 0; i < commas.length; i++) {
                for (int j = 0; j < n; j++) {
                    columns[i][j] = lines[j].length() - lines[j].replace(commas[i], EMPTY).length();
                }
            }

            // Find the most comma character
            int[] nc = new int[commas.length];
            for (int i = 0; i < columns.length; i++) {
                Map<Integer, Integer> c = new HashMap<>();
                for (int j : columns[i]) {
                    if (j == 0) continue;
                    Integer co = c.get(j);
                    c.put(j, co != null ? co + 1 : 1);
                }
                if (c.isEmpty()) continue;

                if (c.size() == 1) {
                    Map.Entry<Integer, Integer> entry = c.entrySet().iterator().next();
                    if (entry.getKey() > 65535) {
                        throw new IOException("Too many columns occur. Max columns 65535 but has " + entry.getKey());
                    }
                    // there only read 10 lines, 4-bits be used
                    nc[i] = (entry.getKey() << 4) + entry.getValue();
                } else {
                    int mv = 0, mk = 0;
                    for (Map.Entry<Integer, Integer> entry : c.entrySet()) {
                        if (entry.getValue() > mv || entry.getValue() == mv && entry.getKey() > mk) {
                            mv = entry.getValue();
                            mk = entry.getKey();
                            nc[i] = (entry.getKey() << 4) + entry.getValue();
                        }
                    }
                }
            }

            O o = new O(0);
            o.line = bom != null ? 1 : 0;
            o.charset = charset;
            n = 0;
            // Find the final comma and column
            for (int i = 0; i < nc.length; i++) {
                int size = nc[i] >>> 4;
                if (size++ == 0) continue;
                int count = nc[i] & 0x0F;
                if (count > n) {
                    n = count;
                    o.offset = size;
                    o.value = commas[i];
                } else if (count == n && size > o.offset) {
                    o.offset = size;
                    o.value = commas[i];
                }
            }

            // Comma character not ',', '\t' or ';'
            if (o.offset == 0) {
                int count = 0;
                for (int c : nc) {
                    count += c;
                }
                // All top 10 row has only one word
                if (count == 0) {
                    o.offset = 1;
                    o.value = commas[0];
                } else {
                    throw new IOException("Unknown comma character, Please specify a separator.");
                }
            }
            return o;
        }
    }

    private static Charset charsetTest(Path path) throws IOException {
        Charset bom = null;
        try (InputStream is = Files.newInputStream(path)) {
            byte[] header = new byte[8];
            int n = is.read(header);
            if (n < 1) return null;
            // 16-bit Unicode
            if (n >= 2) {
                // little-endian byte order
                if ((header[0] & 0xFF) == 0xFF && (header[1] & 0xFF) == 0xFE) {
                    bom = StandardCharsets.UTF_16LE; // UTF-16/UCS-2
                    // 32-bit Unicode
                    if (n >= 4 && header[2] == 0x0 && header[3] == 0x0) {
                        bom = Charset.forName("UTF-32LE"); // UTF-32/UCS-4
                    }
                }
                // big-endian byte order
                else if ((header[0] & 0xFF) == 0xFE && (header[1] & 0xFF) == 0xFF) {
                    bom = StandardCharsets.UTF_16BE;
                }
            }
            // 8-bit Unicode
            if (n >= 3 && (header[0] & 0xFF) == 0xEF && (header[1] & 0xFF) == 0xBB && (header[2] & 0xFF) == 0xBF) {
                bom = StandardCharsets.UTF_8; // UTF-8
            }
            // big-endian byte order UTF-32/UCS-4
            if (n >= 4 && (header[0] & 0xFF) == 0x0 && (header[1] & 0xFF) == 0x0
                && (header[2] & 0xFF) == 0xFE && (header[3] & 0xFF) == 0xFF) {
                bom = Charset.forName("UTF-32BE");
            }
        }
        return bom;
    }

    /**
     * Parse char array
     *
     * @param chars the data array
     * @param len the size of data
     * @param o a cache object
     * @param comma the separate character
     * @return true if the char array has more integral string
     */
    private static boolean parse(char[] chars, int len, O o, char comma) {
        int offset = o.offset, i = offset, iq = -1, qn = 0;
        // the first character is '"'
        boolean quoted = chars[i] == QUOTE
            // an integral string
            , integral = false
            // if data size less than block length
            , last_block = len < chars.length;

        if (quoted) i++;
        for (; i < len; i++) {
            char c = chars[i];
            if (c == QUOTE) {
                if (i >= len - 1) {
                    integral = true;
                    i++;
                    break;
                }
                if (quoted) {
                    if (chars[i + 1] == QUOTE) {
                        iq = ++i;
                        continue;
                    } else if (chars[i + 1] != comma && chars[i + 1] != LF && !(chars[i + 1] == CR && i < len - 2 && chars[i + 2] == LF)) {
                        throw new RuntimeException("line-number: " + o.line + " (zero-base). Comma-separated values" +
                            " format error.\nInvalid char between encapsulated token and delimiter.");
                    }
                } else qn++;
                if (!last_block && chars[i + 1] != comma && chars[i + 1] != LF
                    && !(chars[i + 1] == CR && i < len - 2 && chars[i + 2] == LF)) {
                    throw new RuntimeException("line-number: " + o.line + " (zero-base). Comma-separated values " +
                        "format error.\nA (double) quote character in a field must be represented by two (double) quote characters.");
                }
                i++;
                if (qn == 0) {
                    integral = true;
                    break;
                }
            } else if (c == comma || c == LF) {
                if (!quoted) {
                    integral = true;
                    break;
                }
            }
        }

        if (!integral && last_block) {
            if (!quoted) integral = true;
            else throw new RuntimeException("line-number: " + o.line + " (zero-base). Comma-separated values " +
                "format error.\nEOF reached before encapsulated token finished.");
        }

        if (integral) {
            if (quoted) offset++;
            // an integral string
            if (offset == i && chars[offset] == LF || offset - i == 1 && chars[offset] == CR && chars[i] == LF)
                o.value = null;
            else {
                o.value = i - offset > 0
                    ? trim(chars, offset, quoted || chars[i - 1] == CR ? i - offset - 1 : i - offset, iq) : EMPTY;
            }
            if (i < len - 1 && chars[i] == CR && chars[i + 1] == LF) {
                i += 1;
                o.newLine = true;
            } else if (i < len) {
                o.newLine = chars[i] == LF;
            } else o.newLine = false;
            offset = i + 1;

            // reset the offset
            o.offset = offset;
        }
        return integral;
    }

    /**
     * Returns a string whose value is this string, with any leading and trailing
     * whitespace removed and double-quoted character convert to single-quoted character.
     *
     * @param chars a block data
     * @param offset initial offset of the block data.
     * @param size length of the integral string
     * @return string
     */
    private static String trim(char[] chars, int offset, int size, int iq) {
        if (size > 0) {
            int len = offset + size;
            if (iq >= 0) {
                System.arraycopy(chars, iq, chars, iq - 1, len - iq);
                len--;
                for (int i = iq - 1; i > offset; i--) {
                    if (chars[i] == QUOTE && chars[i - 1] == QUOTE) {
                        System.arraycopy(chars, i, chars, i - 1, len - i);
                        i--;
                        len--;
                    }
                }
            }
            return new String(chars, offset, len - offset);
        }
        return EMPTY;
    }

    // --- PUBLIC inner Writer

    /**
     * A CSV format file writer.
     */
    public static class Writer implements Closeable {

        private final BufferedWriter writer;
        // Comma separator character, default ','
        private char separator = COMMA;
        private int column;
        // The column index
        private int i;
        private char[] cb;
        private int offset;
        private final static int length = 8192;

        /**
         * Line separator string.  This is the value of the line.separator
         * property at the moment that the stream was created.
         */
        private final char[] lineSeparator = System.lineSeparator().toCharArray();

        /**
         * Create a CSV format writer
         *
         * @param path the storage path
         * @throws IOException If an I/O error occurs
         */
        private Writer(Path path) throws IOException {
            this(path, StandardCharsets.UTF_8);
        }

        /**
         * Create a CSV format writer
         *
         * @param path the storage path
         * @param charset the charset to use for encoding
         * @throws IOException If an I/O error occurs
         */
        private Writer(Path path, Charset charset) throws IOException {
            this.writer = Files.newBufferedWriter(path, charset);
            init();
        }

        /**
         * Create a CSV format writer
         *
         * @param writer the output
         */
        private Writer(BufferedWriter writer) {
            this.writer = writer;
            init();
        }

        private void init() {
            cb = new char[length];
        }

        /**
         * Write csv bytes with BOM
         *
         * <p>Note: This property must be set before writing, otherwise it will be ignored</p>
         *
         * @return current writer
         */
        public Writer writeWithBOM() {
            if (offset == 0 && i == 0 && column == 0) {
                // Write UTF BOM
                cb[offset++] = '\uFEFF';
            }
            return this;
        }

        /**
         * Writes a single character
         *
         * @param c int specifying a character to be written
         * @throws IOException If an I/O error occurs
         */
        public void writeChar(char c) throws IOException {
            test();
            if (c == QUOTE) {
                checkBound(4);
                cb[offset++] = QUOTE;
                cb[offset++] = QUOTE;
                cb[offset++] = QUOTE;
                cb[offset++] = QUOTE;
            } else if (c == LF || c == HT || c == separator || c == COMMA) {
                checkBound(3);
                cb[offset++] = QUOTE;
                cb[offset++] = c;
                cb[offset++] = QUOTE;
            } else {
                checkBound(1);
                cb[offset++] = c;
            }
        }

        /**
         * Writes a boolean value, will be convert to boolean string upper case
         *
         * @param b the int value to be written
         * @throws IOException If an I/O error occurs
         */
        public void write(boolean b) throws IOException {
            test();
            if (b) {
                checkBound(4);
                cb[offset++] = 'T';
                cb[offset++] = 'R';
                cb[offset++] = 'U';
                cb[offset++] = 'E';
            } else {
                checkBound(5);
                cb[offset++] = 'F';
                cb[offset++] = 'A';
                cb[offset++] = 'L';
                cb[offset++] = 'S';
                cb[offset++] = 'E';
            }
        }

        /**
         * Writes a int value, will be convert to int string
         *
         * @param n the int value to be written
         * @throws IOException If an I/O error occurs
         */
        public void write(int n) throws IOException {
            test();
            toChars(n);
        }

        /**
         * Writes a long value, will be convert to long string
         *
         * @param l the long value to be written
         * @throws IOException If an I/O error occurs
         */
        public void write(long l) throws IOException {
            test();
            toChars(l);
        }

        /**
         * Writes a single-precision floating-point value, will be convert to single-precision string
         *
         * @param f the single-precision floating-point to be written
         * @throws IOException If an I/O error occurs
         */
        public void write(float f) throws IOException {
            test();
            String fs = Float.toString(f);
            int len = fs.length();
            checkBound(len);
            fs.getChars(0, len, cb, offset);
            offset += len;
        }

        /**
         * Writes a double-precision floating-point value, will be convert to double-precision string
         *
         * @param d the double-precision floating-point to be written
         * @throws IOException If an I/O error occurs
         */
        public void write(double d) throws IOException {
            test();
            String ds = Double.toString(d);
            int len = ds.length();
            checkBound(len);
            ds.getChars(0, len, cb, offset);
            offset += len;
        }

        /**
         * Compression and escape char sequence
         * - line-break, double-quote or commas should be quoted.
         * - A (double) quote character in a field must be represented by two (double) quote characters.
         *
         * @param text the string to be written
         * @throws IOException If an I/O error occurs
         */
        public void write(String text) throws IOException {
            if (text != null && !text.isEmpty()) write(text.toCharArray());
            else writeEmpty();
        }

        /**
         * Compression and escape char sequence
         * - line-break, double-quote or commas should be quoted.
         * - A (double) quote character in a field must be represented by two (double) quote characters.
         *
         * @param chars the char sequence to be written
         * @throws IOException If an I/O error occurs
         */
        public void write(char[] chars) throws IOException {
            write(chars, 0, chars.length);
        }

        /**
         * Compression and escape char sequence
         * - line-break, double-quote or commas should be quoted.
         * - A (double) quote character in a field must be represented by two (double) quote characters.
         *
         * @param chars the char sequence to be written
         * @param offset the offset index
         * @param size size of characters
         * @throws IOException If an I/O error occurs
         */
        public void write(char[] chars, int offset, int size) throws IOException {
            test();
            int i = 0;
            int last = offset;
            boolean quoted = false, shouldBeQuoted = false;

            for ( ; i < size; ) {
                char c = chars[i++];

                // A (double) quote character in a field must be represented
                // by two (double) quote characters.
                if (c == QUOTE) {
                    quoted = true;
                    if (last == offset) {
                        checkBound(1);
                        cb[this.offset++] = QUOTE;
                    }
                    checkBound(i - last + 1);
                    System.arraycopy(chars, last, cb, this.offset, i - last);
                    this.offset += (i - last);
                    cb[this.offset++] = QUOTE;
                    last = i;
                }

                else if (c == LF || c == HT || c == separator || c == COMMA) {
                    shouldBeQuoted = true;
                }

            }

            if (quoted) {
                checkBound(i - last + 1);
                System.arraycopy(chars, last, cb, this.offset, i - last);
                this.offset += (i - last);
                cb[this.offset++] = QUOTE;
            }
            else if (shouldBeQuoted) {
                checkBound(size + 2);
                cb[this.offset++] = QUOTE;
                System.arraycopy(chars, offset, cb, this.offset, size);
                this.offset += size;
                cb[this.offset++] = QUOTE;
            }
            else {
                checkBound(size);
                System.arraycopy(chars, offset, cb, this.offset, size);
                this.offset += size;
            }

        }

        /**
         * Writes a empty column.
         * @throws IOException If an I/O error occurs
         */
        public void writeEmpty() throws IOException {
            test();
        }

        /**
         * Writes a line separator.  The line separator string is defined by the
         * system property <tt>line.separator</tt>, and is not necessarily a single
         * newline ('\n') character.
         *
         * @exception  IOException  If an I/O error occurs
         */
        public void newLine() throws IOException {
            checkBound(lineSeparator.length);
            System.arraycopy(lineSeparator, 0, cb, offset, lineSeparator.length);
            offset += lineSeparator.length;
            if (column == 0) column = i;
            i = 0;
        }

        /**
         * Test the column index
         *
         * @return true if first column
         */
        private boolean test() throws IOException {
            boolean first = i == 0;
            i++;
            if (column > 0 && i > column) {
                // FIXME maybe throw an exception
                LOGGER.warn("Each record should contain the same number of comma-separated fields.");
            }
            if (!first) {
                checkBound(1);
                cb[offset++] = separator;
            }
            return first;
        }

        private void flush() throws IOException {
            writer.write(cb, 0, offset);
            offset = 0;
        }

        private void checkBound(int size) throws IOException {
            if (offset + size > length) {
                flush();
            }
        }

        private void toChars(int i) throws IOException {
            if (i == Integer.MIN_VALUE) {
                checkBound(MIN_INTEGER_CHARS.length);
                System.arraycopy(MIN_INTEGER_CHARS, 0, cb, offset, MIN_INTEGER_CHARS.length);
                offset += MIN_INTEGER_CHARS.length;
            } else {
                int size = stringSize(i);
                checkBound(size);
                getChars(i, offset += size, cb);
            }
        }

        private void toChars(long i) throws IOException {
            if (i == Long.MIN_VALUE) {
                checkBound(MIN_LONG_CHARS.length);
                System.arraycopy(MIN_LONG_CHARS, 0, cb, offset, MIN_LONG_CHARS.length);
                offset += MIN_LONG_CHARS.length;
            } else {
                int size = stringSize(i);
                checkBound(size);
                getChars(i, offset += size, cb);
            }
        }

        @Override
        public void close() throws IOException {
            if (writer != null) {
                if (offset > 0) {
                    flush();
                }
                writer.close();
            }
        }
    }
}

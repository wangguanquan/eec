package net.cua.excel.reader;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.UncheckedIOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Stream;
import java.util.stream.StreamSupport;

/**
 * Create by guanquan.wang at 2018-09-22
 */
public class Sheet implements AutoCloseable {
    Logger logger = LogManager.getLogger(getClass());
    Sheet() {}

    private String name;
    private int index // per sheet index of workbook
            , size = -1; // size of rows per sheet
    private Path path;
    private SharedString sst;
    private int startRow = -1; // row index of data
    private Row header;

    public void setName(String name) {
        this.name = name;
    }

    public void setPath(Path path) {
        this.path = path;
    }

    public void setSst(SharedString sst) {
        this.sst = sst;
    }

    public String getName() {
        return name;
    }

    public Path getPath() {
        return path;
    }

    public int getIndex() {
        return index;
    }

    void setIndex(int index) {
        this.index = index;
    }

    /**
     * size of sheet. -1: unknown size
     * @return number of rows
     */
    public int getSize() {
        return size;
    }

    public int getStartRow() {
        return startRow;
    }

    /**
     * this.rows()方法第一行就是头部信息，
     * rowsWithOutHeader()方法会跳过头部信息，此时可以使用此方法获得头部信息。
     * @return HeaderRow
     */
    public Row getHeader() {
        if (header == null) {
            Row row = findRow0();
            if (row != null)
                header = row.asHeader();
        }
        return header;
    }

    @Override
    public String toString() {
        return name + " : " + path;
    }

    /////////////////////////////////Read sheet file/////////////////////////////////
    private BufferedReader reader;
    private char[] cb; // buffer
    private int nChar, length;
    private boolean eof = false;

    private Row sRow;

    /**
     * load sheet.xml as BufferedReader
     * @return
     * @throws IOException
     */
    Sheet load() throws IOException {
        logger.debug("load {}", path.toString());
        reader = Files.newBufferedReader(path);
        cb = new char[8192];
        nChar = 0;
        loopA: for ( ; ; ) {
            length = reader.read(cb);
            if (length < 11) break;
            // read size
            if (nChar == 0) {
                String line = new String(cb, 0, 512);
                String size = "<dimension ref=\"";
                int index = line.indexOf(size), end = index > 0 ? line.indexOf('"', index+=(size.length()+1)) : -1;
                if (end > 0) {
                    String __ = line.substring(index, end);
                    Pattern pat = Pattern.compile("[A-Z]+(\\d+):[A-Z]+(\\d+)");
                    Matcher mat = pat.matcher(__);
                    if (mat.matches()) {
                        int from = Integer.parseInt(mat.group(1)), to = Integer.parseInt(mat.group(2));
                        this.startRow = from;
                        this.size = to - from;
                    } else {
                        pat = Pattern.compile("[A-Z]+(\\d+)");
                        mat = pat.matcher(__);
                        if (mat.matches()) {
                            this.startRow = Integer.parseInt(mat.group(1));
                        }
                    }
                    nChar += end;
                }
            }
            // find index of <sheetData>
            for (; nChar < length - 11; nChar++) {
                if (cb[nChar] == '<' && cb[nChar+1] == 's' && cb[nChar+2] == 'h'
                        && cb[nChar+3] == 'e' && cb[nChar+4] == 'e' && cb[nChar+5] == 't'
                        && cb[nChar+6] == 'D' && cb[nChar+7] == 'a' && cb[nChar+8] == 't'
                        && cb[nChar+9] == 'a' && cb[nChar+10] == '>') {
                    nChar += 11;
                    break loopA;
                }
            }
        }
        sRow = new Row(sst); // share row space

        return this;
    }

    /**
     * iterator rows
     * @return
     * @throws IOException
     */
    public Row nextRow() throws IOException {
        if (eof) return null;
        boolean endTag = false;
        int start = nChar;
        // find end of row tag
        for (; nChar < length - 6; nChar++) {
            if (cb[nChar] == '<' && cb[nChar+1] == '/' && cb[nChar+2] == 'r'
                    && cb[nChar+3] == 'o' && cb[nChar+4] == 'w' && cb[nChar+5] == '>') {
                nChar += 6;
                endTag = true;
                break;
            }
        }

        /* Load more when not found end of row tag */
        if (!endTag) {
            int n;
            System.arraycopy(cb, start, cb, 0, n = length - start);
            length = reader.read(cb, n, cb.length - n);
            // end of file
            if (length < 0) {
                eof = true;
                reader.close(); // close reader
                reader = null; // wait GC
                logger.debug("end of file.");
                return null;
            }
            nChar = 0;
            length += n;
            return nextRow();
        }

        // share row
        return sRow.with(cb, start, nChar - start);
    }

    protected Row findRow0() {
        char[] cb = new char[8192];
        int nChar = 0, length;
        // reload file
        try (BufferedReader reader = Files.newBufferedReader(path)) {
            loopA:
            for (; ; ) {
                length = reader.read(cb);
                if (length < 11) break;
                // find index of <sheetData>
                for (; nChar < length - 11; nChar++) {
                    if (cb[nChar] == '<' && cb[nChar + 1] == 's' && cb[nChar + 2] == 'h'
                            && cb[nChar + 3] == 'e' && cb[nChar + 4] == 'e' && cb[nChar + 5] == 't'
                            && cb[nChar + 6] == 'D' && cb[nChar + 7] == 'a' && cb[nChar + 8] == 't'
                            && cb[nChar + 9] == 'a' && cb[nChar + 10] == '>') {
                        nChar += 11;
                        break loopA;
                    }
                }
            }
        } catch (IOException e) {
            logger.error("Read header row error.");
            return null;
        }

        boolean eof = false;
        if (eof) return null;
        boolean endTag = false;
        int start = nChar;
        // find end of row tag
        for (; nChar < length - 6; nChar++) {
            if (cb[nChar] == '<' && cb[nChar+1] == '/' && cb[nChar+2] == 'r'
                    && cb[nChar+3] == 'o' && cb[nChar+4] == 'w' && cb[nChar+5] == '>') {
                nChar += 6;
                endTag = true;
                break;
            }
        }

        if (!endTag) {
            // too big
            return null;
        }

        // row
        return new Row(sst).with(cb, start, nChar - start);
    }

    /**
     *
     * @return
     */
    public Iterator<Row> iterator() {
        return iter;
    }

    // iterator foreach rows
    private Iterator<Row> iter = new Iterator<Row>() {
        Row nextRow = null;

        @Override
        public boolean hasNext() {
            if (nextRow != null) {
                return true;
            } else {
                try {
                    nextRow = nextRow();
                    return (nextRow != null);
                } catch (IOException e) {
                    throw new UncheckedIOException(e);
                }
            }
        }

        @Override
        public Row next() {
            if (nextRow != null || hasNext()) {
                Row next = nextRow;
                nextRow = null;
                return next;
            } else {
                throw new NoSuchElementException();
            }
        }
    };

    /**
     * stream of all rows
     * @return a {@code Stream<Row>} providing the lines of row
     *         described by this {@code Sheet}
     * @since 1.8
     */
    public Stream<Row> rows() {
        return StreamSupport.stream(Spliterators.spliteratorUnknownSize(
                iter, Spliterator.ORDERED | Spliterator.NONNULL), false);
    }

    /**
     * stream with out header row
     * @return a {@code Stream<Row>} providing the lines of row
     *         described by this {@code Sheet}
     * @since 1.8
     */
    public Stream<Row> rowsWithOutHeader() {
        if (iter.hasNext()) {
            this.header = iter.next().asHeader(); // skip first row
        }
        return StreamSupport.stream(Spliterators.spliteratorUnknownSize(
                iter, Spliterator.ORDERED | Spliterator.NONNULL), false);
    }

    /**
     * close reader
     * @throws IOException
     */
    public void close() throws IOException {
        cb = null;
        if (reader != null) {
            reader.close();
        }
    }
}

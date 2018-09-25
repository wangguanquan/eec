package net.cua.export.reader;

import java.io.IOException;
import java.io.UncheckedIOException;
import java.util.*;
import java.util.stream.Stream;
import java.util.stream.StreamSupport;

/**
 * Create by guanquan.wang at 2018-09-22
 */
public class Sheet {
    String name; // sheet name
    int rows; // number of rows
    int cursor; // index of next element to return

    public Row readLine() throws IOException {
        // TODO
        return null;
    }
//    Iterator<Row> iterator

    private class Itr implements Iterator<Row> {
        int cursor;       // index of next element to return
        int lastRet = -1; // index of last element returned; -1 if no such

        public boolean hasNext() {
            return cursor != rows;
        }

        public Row next() {
            int i = cursor;
            if (i >= rows)
                throw new NoSuchElementException();
            cursor = i + 1;
            Row e;
            try {
                e = readLine();
            } catch (IOException ex) {
                throw new UncheckedIOException(ex);
            }
            return e;
        }
    }
    /**
     * to streams
     * @return sheet stream
     */
    public Stream<Row> rows() {
        Iterator<Row> ite = new Iterator<Row>() {
            Row nextRow = null;
            @Override
            public boolean hasNext() {
                if (nextRow != null) {
                    return true;
                } else {
                    try {
                        nextRow = readLine();
                        return (nextRow != null);
                    } catch (IOException e) {
                        throw new UncheckedIOException(e);
                    }
                }
            }

            @Override
            public Row next() {
                if (nextRow != null || hasNext()) {
                    Row e = nextRow;
                    nextRow = null;
                    return e;
                } else {
                    throw new NoSuchElementException();
                }
            }
        };
        return StreamSupport.stream(Spliterators.spliterator(ite
                , rows, Spliterator.ORDERED | Spliterator.NONNULL), false);
    }
}

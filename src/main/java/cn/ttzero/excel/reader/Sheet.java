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

import cn.ttzero.excel.util.StringUtil;

import java.io.IOException;
import java.util.Iterator;
import java.util.Spliterator;
import java.util.Spliterators;
import java.util.stream.Stream;
import java.util.stream.StreamSupport;

/**
 * Create by guanquan.wang at 2019-04-17 11:36
 */
public interface Sheet extends AutoCloseable {

    /**
     * The worksheet name
     *
     * @return the sheet name
     */
    String getName();

    /**
     * The index of worksheet located at the workbook
     *
     * @return the index(zero base)
     */
    int getIndex();

    /**
     * size of rows.
     *
     * @return size of rows
     * -1: unknown size
     */
    int getSize();

    /**
     * Test Worksheet is hidden
     */
    boolean isHidden();

    /**
     * Test Worksheet is show
     */
    default boolean isShow() {
        return !isHidden();
    }

    /**
     * Returns the header of the list.
     * The first non-empty line defaults to the header information.
     *
     * @return the HeaderRow
     */
    Row getHeader();

    /**
     * Set the binding type
     *
     * @param clazz the binding type
     * @return sheet
     */
    Sheet bind(Class<?> clazz);

    /**
     * Load the sheet data
     *
     * @return sheet
     * @throws IOException if io error occur
     */
    Sheet load() throws IOException;

    /**
     * Iterating each row of data contains header information and blank lines
     *
     * @return a row iterator
     */
    Iterator<Row> iterator();

    /**
     * Iterating over data rows without header information and blank lines
     *
     * @return a row iterator
     */
    Iterator<Row> dataIterator();

    /**
     * Reset the {@link Sheet}'s row index to begging
     *
     * @return the unread {@link Sheet}
     */
    default Sheet reset() throws IOException {
        throw new UnsupportedOperationException();
    }

    /**
     * Return a stream of all rows
     *
     * @return a {@code Stream<Row>} providing the lines of row
     * described by this {@code Sheet}
     * @since 1.8
     */
    default Stream<Row> rows() {
        return StreamSupport.stream(Spliterators.spliteratorUnknownSize(
            iterator(), Spliterator.ORDERED | Spliterator.NONNULL), false);
    }

    /**
     * Return stream with out header row and empty rows
     *
     * @return a {@code Stream<Row>} providing the lines of row
     * described by this {@code Sheet}
     * @since 1.8
     */
    default Stream<Row> dataRows() {
        return StreamSupport.stream(Spliterators.spliteratorUnknownSize(
            dataIterator(), Spliterator.ORDERED | Spliterator.NONNULL), false);
    }

    /**
     * Convert column mark to int
     *
     * @param col column mark
     * @return int value
     */
    default int col2Int(String col) {
        if (StringUtil.isEmpty(col)) return 1;
        char[] values = col.toCharArray();
        int n = 0;
        for (char value : values) {
            if (value < 'A' || value > 'Z')
                throw new ExcelReadException("Column mark out of range: " + col);
            n = n * 26 + value - 'A' + 1;
        }
        return n;
    }

    /**
     * Close resource
     *
     * @throws IOException if io error occur
     */
    void close() throws IOException;
}

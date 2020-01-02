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

package org.ttzero.excel.reader;

import org.ttzero.excel.manager.Const;
import org.ttzero.excel.util.CSVUtil;
import org.ttzero.excel.util.StringUtil;

import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Iterator;
import java.util.Spliterator;
import java.util.Spliterators;
import java.util.stream.Stream;
import java.util.stream.StreamSupport;

import static org.ttzero.excel.reader.Cell.BOOL;
import static org.ttzero.excel.reader.Cell.DOUBLE;
import static org.ttzero.excel.reader.Cell.FUNCTION;
import static org.ttzero.excel.reader.Cell.INLINESTR;
import static org.ttzero.excel.reader.Cell.LONG;
import static org.ttzero.excel.reader.Cell.NUMERIC;
import static org.ttzero.excel.reader.Cell.SST;
import static org.ttzero.excel.util.DateUtil.toLocalDate;
import static org.ttzero.excel.util.DateUtil.toTimestamp;

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
     *      -1: unknown size
     * @deprecated use {@link #getDimension()} to getting full range address
     */
    @Deprecated
    int getSize();

    /**
     * Returns The range address of the used area in
     * the current sheet
     * <p>
     * NOTE: This method can only guarantee accurate row ranges
     *
     * @return worksheet {@link Dimension} ranges
     */
    Dimension getDimension();

    /**
     * Test Worksheet is hidden
     *
     * @return true if current worksheet is hidden
     */
    boolean isHidden();

    /**
     * Test Worksheet is show
     *
     * @return true if current worksheet is show
     */
    default boolean isShow() {
        return !isHidden();
    }

    /**
     * Returns the header of the list.
     * The first non-empty line defaults to the header information.
     *
     * @return the {@link HeaderRow}
     */
    Row getHeader();

    /**
     * Set the binding type
     *
     * @param clazz the binding type
     * @return the {@link Sheet}
     */
    Sheet bind(Class<?> clazz);

    /**
     * Load the sheet data
     *
     * @return the {@link Sheet}
     * @throws IOException if I/O error occur
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
     * @throws ExcelReadException if I/O error occur.
     */
    default Sheet reset() {
        throw new UnsupportedOperationException();
    }

    /**
     * Return a stream of all rows
     *
     * @return a {@code Stream&lt;Row&gt;} providing the lines of row
     * described by this {@link Sheet}
     * @since 1.8
     */
    default Stream<Row> rows() {
        return StreamSupport.stream(Spliterators.spliteratorUnknownSize(
            iterator(), Spliterator.ORDERED | Spliterator.NONNULL), false);
    }

    /**
     * Return stream with out header row and empty rows
     *
     * @return a {@code Stream&lt;Row&gt;} providing the lines of row
     * described by this {@link Sheet}
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
     * @throws IOException if I/O error occur
     */
    void close() throws IOException;

    /**
     * Save file as Comma-Separated Values. Each worksheet corresponds to
     * a csv file. Default charset is 'UTF8' and separator character is ','.
     * @param path the output storage path
     */
    default void saveAsCSV(Path path) throws IOException {
        // Create path if not exists
        if (!Files.exists(path)) {
            Files.createDirectories(path);
        }
        if (Files.isDirectory(path)) {
            path = path.resolve(getName() + Const.Suffix.CSV);
        }

        saveAsCSV(Files.newOutputStream(path));
    }

    /**
     * Save file as Comma-Separated Values. Each worksheet corresponds to
     * a csv file. Default charset is 'UTF8' and separator character is ','.
     * @param os the output
     */
    default void saveAsCSV(OutputStream os) throws IOException {
        try (CSVUtil.Writer writer = CSVUtil.newWriter(os)) {
            for (Iterator<Row> iter = iterator(); iter.hasNext(); ) {
                Row row = iter.next();
                if (row.isEmpty()) continue;
                for (int i = row.fc; i < row.lc; i++) {
                    Cell c = row.cells[i];
                    switch (c.t) {
                        case SST:
                            if (c.sv == null) {
                                c.setSv(row.sst.get(c.nv));
                            }
                            writer.write(c.sv);
                            break;
                        case INLINESTR:
                        case FUNCTION: // Formula string value
                            writer.write(c.sv);
                            break;
                        case BOOL:
                            writer.write(c.bv);
                            break;
                        case NUMERIC:
                            if (!row.styles.fastTestDateFmt(c.s)) writer.write(c.nv);
                            else writer.write(toLocalDate(c.nv).toString());
                            break;
                        case LONG:
                            writer.write(c.lv);
                            break;
                        case DOUBLE:
                            if (!row.styles.fastTestDateFmt(c.s)) writer.write(c.dv);
                            else writer.write(toTimestamp(c.dv).toString());
                            break;
                        default:
                            writer.writeEmpty();
                    }
                }
                writer.newLine();
            }
        }
    }

    /**
     * Make reader parse the formula
     */
    void parseFormula();
}

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

package org.ttzero.excel.entity;

import org.ttzero.excel.manager.Const;
import org.ttzero.excel.reader.Cell;
import org.ttzero.excel.util.CSVUtil;
import org.ttzero.excel.util.FileUtil;

import java.io.BufferedWriter;
import java.io.IOException;
import java.io.InputStream;
import java.io.Reader;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;

/**
 * Create by guanquan.wang at 2019-09-26 08:33
 */
public class CSVSheet extends Sheet {

    // The csv source file path
    private Path path;
    private CSVUtil.RowsIterator iterator;
    private boolean shouldClean;
    private boolean hasHeader;

    /**
     * Constructor worksheet
     */
    public CSVSheet() {
        super();
    }

    /**
     * Constructor worksheet
     * @param path the csv source file path
     */
    public CSVSheet(Path path) {
        this.path = path;
    }

    /**
     * Constructor worksheet
     * @param name the worksheet name
     * @param path the csv source file path
     */
    public CSVSheet(String name, Path path) {
        super(name);
        this.path = path;
    }

    /**
     * Constructor worksheet
     * @param is the csv source InputStream
     * @throws IOException if I/O error occur.
     */
    public CSVSheet(InputStream is) throws IOException {
        this(null, is);
    }

    /**
     * Constructor worksheet
     * @param name the worksheet name
     * @param is the csv source InputStream
     * @throws IOException if I/O error occur.
     */
    public CSVSheet(String name, InputStream is) throws IOException {
        super(name);
        path = Files.createTempFile(Const.EEC_PREFIX, String.valueOf(id));
        Files.copy(is, path, StandardCopyOption.REPLACE_EXISTING);
        shouldClean = true;
    }

    /**
     * Constructor worksheet
     * @param reader the csv source InputStream
     * @throws IOException if I/O error occur.
     */
    public CSVSheet(Reader reader) throws IOException {
        this(null, reader);
    }

    /**
     * Constructor worksheet
     * @param name the worksheet name
     * @param reader the csv source InputStream
     * @throws IOException if I/O error occur.
     */
    public CSVSheet(String name, Reader reader) throws IOException {
        super(name);
        path = Files.createTempFile(Const.EEC_PREFIX, String.valueOf(id));
        shouldClean = true;
        char[] chars = new char[8192];
        int n;
        try (BufferedWriter bw = Files.newBufferedWriter(path)) {
            while ((n = reader.read(chars)) > 0) {
                bw.write(chars, 0, n);
            }
        }
    }

    /**
     * Setting header flag
     *
     * @param hasHeader boolean value
     * @return {@link CSVSheet}
     */
    public CSVSheet setHasHeader(boolean hasHeader) {
        this.hasHeader = hasHeader;
        return this;
    }

    /**
     * Release resources
     *
     * @throws IOException if I/O error occur
     */
    @Override
    public void close() throws IOException {
        iterator.close();
        if (shouldClean) {
            FileUtil.rm_rf(path);
        }
        super.close();
    }

    // Create CSV iterator
    private void init() throws IOException {
        assert path != null && Files.exists(path);
        iterator = CSVUtil.newReader(path).sharedIterator();
    }

    /**
     * Reset the row-block data
     */
    @Override
    protected void resetBlockData() {
        int len = columns.length, n = 0, limit = sheetWriter.getRowLimit() - 1;
        for (int rbs = getRowBlockSize(); n++ < rbs && rows < limit && iterator.hasNext(); rows++) {
            Row row = rowBlock.next();
            row.index = rows;
            Cell[] cells = row.realloc(len);
            String[] csvRow = iterator.next();
            for (int i = 0; i < len; i++) {
                Column hc = columns[i];

                // clear cells
                Cell cell = cells[i];
                cell.clear();

                cell.setSv(csvRow[i]);
                cell.xf = cellValueAndStyle.getStyleIndex(rows, hc, csvRow[i]);
//                cellValueAndStyle.reset(rows, cell, csvRow[i], hc);
            }
        }

        // Paging
        if (rows >= limit) {
            shouldClose = false;
            CSVSheet copy = getClass().cast(clone());
            workbook.insertSheet(id, copy);
        } else shouldClose = true;
    }

    @Override
    public Column[] getHeaderColumns() {
        if (headerReady) return columns;
        try {
            // Create CSV iterator
            init();
            if (!iterator.hasNext()) return columns;
            String[] rows = iterator.next();
            columns = new Column[rows.length];
            for (int i = 0; i < rows.length; i++) {
                // FIXME the column type
                columns[i] = new Column(hasHeader ? rows[i] : null, String.class);
                columns[i].styles = workbook.getStyles();
            }
            headerReady = true;
            if (hasNonHeader()) {
                ((CSVUtil.SharedRowsIterator) iterator).retain();
            }
        } catch (IOException e) {
            throw new ExcelWriteException(e);
        }
        return columns;
    }

    /**
     * Check empty header row
     *
     * @return true if none header row
     */
    @Override
    public boolean hasNonHeader() {
        if (!hasHeader) {
            hasHeader = !super.hasNonHeader();
        }
        return !hasHeader;
    }
}

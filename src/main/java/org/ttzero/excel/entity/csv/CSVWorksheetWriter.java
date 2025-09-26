/*
 * Copyright (c) 2017-2019, guanquan.wang@hotmail.com All Rights Reserved.
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

package org.ttzero.excel.entity.csv;

import org.ttzero.excel.entity.Column;
import org.ttzero.excel.entity.ExcelWriteException;
import org.ttzero.excel.entity.ICellValueAndStyle;
import org.ttzero.excel.entity.IWorksheetWriter;
import org.ttzero.excel.entity.Row;
import org.ttzero.excel.entity.RowBlock;
import org.ttzero.excel.entity.Sheet;
import org.ttzero.excel.manager.Const;
import org.ttzero.excel.reader.Cell;
import org.ttzero.excel.util.CSVUtil;

import java.io.IOException;
import java.nio.charset.Charset;
import java.nio.file.Path;
import java.util.function.BiConsumer;

import static org.ttzero.excel.reader.Cell.BOOL;
import static org.ttzero.excel.reader.Cell.CHARACTER;
import static org.ttzero.excel.reader.Cell.DATE;
import static org.ttzero.excel.reader.Cell.DATETIME;
import static org.ttzero.excel.reader.Cell.DECIMAL;
import static org.ttzero.excel.reader.Cell.DOUBLE;
import static org.ttzero.excel.reader.Cell.INLINESTR;
import static org.ttzero.excel.reader.Cell.LONG;
import static org.ttzero.excel.reader.Cell.NUMERIC;
import static org.ttzero.excel.reader.Cell.SST;
import static org.ttzero.excel.reader.Cell.TIME;
import static org.ttzero.excel.util.DateUtil.toDate;
import static org.ttzero.excel.util.DateUtil.toDateString;
import static org.ttzero.excel.util.DateUtil.toDateTimeString;
import static org.ttzero.excel.util.DateUtil.toTimeString;
import static org.ttzero.excel.util.StringUtil.isNotEmpty;

/**
 * @author guanquan.wang at 2019-08-21 22:19
 */
public class CSVWorksheetWriter implements IWorksheetWriter {
    protected Sheet sheet;
    protected CSVUtil.Writer writer;
    /**
     * A progress window
     */
    protected BiConsumer<Sheet, Integer> progressConsumer;
    // Write BOM header
    protected boolean withBom, ready;
    // Charset 默认UTF-8
    protected Charset charset;
    /**
     * Delimiter char
     */
    protected char delimiter = ',';
    /**
     * 临时路径
     */
    protected Path workSheetPath;

    public CSVWorksheetWriter() { }

    public CSVWorksheetWriter(Sheet sheet) {
        this.sheet = sheet;
    }

    public CSVWorksheetWriter(Sheet sheet, boolean withBom) {
        this.sheet = sheet;
        this.withBom = withBom;
    }

    /**
     * The row limit
     *
     * @return the const value {@code (1 << 31) - 1}
     */
    @Override
    public int getRowLimit() {
        return Integer.MAX_VALUE;
    }

    /**
     * The column limit
     *
     * @return the const value 16_384
     */
    @Override
    public int getColumnLimit() {
        return Const.Limit.MAX_COLUMNS_ON_SHEET;
    }

    /**
     * Settings delimiter char
     *
     * @param delimiter delimiter char
     * @return current CSVWorksheetWriter
     */
    public CSVWorksheetWriter setDelimiter(char delimiter) {
        this.delimiter = delimiter;
        return this;
    }

    @Override
    public IWorksheetWriter setWorksheet(Sheet sheet) {
        this.sheet = sheet;
        return this;
    }

    @Override
    public IWorksheetWriter clone() {
        throw new ExcelWriteException("Overflow the max row limit.");
    }

    /**
     * Returns the worksheet name
     *
     * @return name of worksheet
     */
    @Override
    public String getFileSuffix() {
        return Const.Suffix.CSV;
    }

    @Override
    public void writeData(RowBlock rowBlock) throws IOException {
        if (!ready) {
            Path path = sheet.getWorkbook().getWorkbookWriter().writeBefore();
            initWriter(path);
            // write before
            writeBefore();
        }

        if (progressConsumer == null) writeRowBlock(rowBlock);
        else writeRowBlockFireProgress(rowBlock);
    }

    @Override
    public void close() throws IOException {
        // Write some final info
        sheet.afterSheetAccess(workSheetPath);
        ready = false;
        if (writer != null) writer.close();
    }

    @Override
    public void writeTo(Path root) throws IOException {
        initWriter(root);
        // Get the first block
        RowBlock rowBlock = sheet.nextBlock();

        // write before
        writeBefore();

        if (rowBlock.hasNext()) {
            if (progressConsumer == null) {
                for (; ; ) {
                    // write row-block data
                    writeRowBlock(rowBlock);
                    // end of row
                    if (rowBlock.isEOF()) break;
                    // Get the next block
                    rowBlock = sheet.nextBlock();
                }
            } else {
                for (; ; ) {
                    // write row-block data and fire progress event
                    writeRowBlockFireProgress(rowBlock);
                    // end of row
                    if (rowBlock.isEOF()) break;
                    // Get the next block
                    rowBlock = sheet.nextBlock();
                }
                if (rowBlock.lastRow() != null) progressConsumer.accept(sheet, rowBlock.lastRow().getIndex());
            }
        }
    }

    protected Path initWriter(Path root) throws IOException {
        this.workSheetPath = root.resolve(sheet.getName() + Const.Suffix.CSV);
        // Already initialized
        if (ready) return workSheetPath;
        if (charset == null) {
            writer = CSVUtil.newWriter(workSheetPath, delimiter);
            if (withBom) writer.writeWithBom();
        } else {
            writer = CSVUtil.newWriter(workSheetPath, delimiter, charset);
            // Write BOM only support utf
            if (withBom && charset.name().toLowerCase().startsWith("utf"))
                writer.writeWithBom();
        }
        // Init progress window
        progressConsumer = sheet.getProgressConsumer();
        ready = true;
        return workSheetPath;
    }

    /**
     * Write worksheet header data
     *
     * @throws IOException if I/O error occur
     */
    protected void writeBefore() throws IOException {
        // The header columns
        Column[] columns = sheet.getAndSortHeaderColumns();
        boolean noneHeader = columns == null || columns.length == 0;

        if (!noneHeader) {
            // Write header
            int subColumnSize = columns[0].subColumnSize();
            Column[][] columnsArray = new Column[columns.length][];
            for (int i = 0; i < columns.length; i++) {
                columnsArray[i] = columns[i].toArray();
            }
            for (int i = subColumnSize - 1; i >= 0; i--) {
                for (int j = 0; j < columns.length; j++) {
                    Column hc = columnsArray[j][i];
                    writer.write(isNotEmpty(hc.getName()) ? hc.getName() : hc.key);
                }
                writer.newLine();
            }
        }
    }

    /**
     * Write a row-block
     *
     * @param rowBlock the row-block
     * @throws IOException if I/O error occur.
     */
    protected void writeRowBlock(RowBlock rowBlock) throws IOException {
        for (; rowBlock.hasNext(); writeRow(rowBlock.next())) ;
    }

    /**
     * Write a row-block and fire progress event
     *
     * @param rowBlock the row-block
     * @throws IOException if I/O error occur.
     */
    protected void writeRowBlockFireProgress(RowBlock rowBlock) throws IOException {
        Row row;
        while (rowBlock.hasNext()) {
            row = rowBlock.next();
            writeRow(row);
            // Fire progress
            if (row.getIndex() % 1_000 == 0) progressConsumer.accept(sheet, row.getIndex());
        }
    }

    /**
     * Write a row data
     *
     * @param row a row data
     * @throws IOException if I/O error occur
     */
    protected void writeRow(Row row) throws IOException {
        Cell[] cells = row.getCells();
        for (Cell cell : cells) {
            switch (cell.t) {
                case INLINESTR:
                case SST      : writer.write(cell.stringVal);                           break;
                case NUMERIC  : writer.write(cell.intVal);                              break;
                case LONG     : writer.write(cell.longVal);                             break;
                case DOUBLE   : writer.write(cell.doubleVal);                           break;
                case BOOL     : writer.write(cell.boolVal);                             break;
                case DECIMAL  : writer.write(cell.decimal.toString());                  break;
                case CHARACTER: writer.writeChar(cell.charVal);                         break;
                case DATE     : writer.write(toDateString(toDate(cell.intVal)));        break;
                case DATETIME : writer.write(toDateTimeString(toDate(cell.doubleVal))); break;
                case TIME     : writer.write(toTimeString(toDate(cell.doubleVal)));     break;
                default       : writer.writeEmpty();
            }
        }
        writer.newLine();
    }

    /**
     * 设置字符集
     *
     * @param charset Charset
     * @return 当前Writer
     */
    public CSVWorksheetWriter setCharset(Charset charset) {
        this.charset = charset;
        return this;
    }

    /**
     * 返回CSV数据样式转换器，该转换器将所有数据转为字符器格式并忽略所有样式
     *
     * @return {@link ICellValueAndStyle}
     */
    @Override
    public ICellValueAndStyle getCellValueAndStyle() {
        return new CSVCellValueAndStyle();
    }
}

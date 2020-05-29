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

package org.ttzero.excel.entity.csv;

import org.ttzero.excel.entity.ExcelWriteException;
import org.ttzero.excel.entity.IWorksheetWriter;
import org.ttzero.excel.entity.Row;
import org.ttzero.excel.entity.RowBlock;
import org.ttzero.excel.entity.Sheet;
import org.ttzero.excel.manager.Const;
import org.ttzero.excel.reader.Cell;
import org.ttzero.excel.util.CSVUtil;
import org.ttzero.excel.util.DateUtil;

import java.io.IOException;
import java.nio.file.Path;
import java.util.function.Supplier;

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
import static org.ttzero.excel.util.StringUtil.isNotEmpty;

/**
 * @author guanquan.wang at 2019-08-21 22:19
 */
public class CSVWorksheetWriter implements IWorksheetWriter {
    private Sheet sheet;
    private CSVUtil.Writer writer;

    public CSVWorksheetWriter(Sheet sheet) {
        this.sheet = sheet;
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

    @Override
    public void writeTo(Path path, Supplier<RowBlock> supplier) throws IOException {
        Path workSheetPath = initWriter(path);
        // Get the first block
        RowBlock rowBlock = supplier.get();

        // write before
        writeBefore();

        if (rowBlock != null && rowBlock.hasNext()) {
            do {
                // write row-block data
                writeRow(rowBlock.next());
                // end of row
                if (rowBlock.isEOF()) break;
            } while ((rowBlock = supplier.get()) != null);
        }
        // Write some final info
        sheet.afterSheetAccess(workSheetPath);
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

    @Override
    public void close() throws IOException {
        if (writer != null)
            writer.close();
    }

    @Override
    public void writeTo(Path root) throws IOException {
        Path workSheetPath = initWriter(root);
        // Get the first block
        RowBlock rowBlock = sheet.nextBlock();

        // write before
        writeBefore();

        if (rowBlock.hasNext()) {
            for (; ; ) {
                // write row-block data
                for (; rowBlock.hasNext(); writeRow(rowBlock.next())) ;
                // end of row
                if (rowBlock.isEOF()) break;
                // Get the next block
                rowBlock = sheet.nextBlock();
            }
        }
        // Write some final info
        sheet.afterSheetAccess(workSheetPath);
    }

    protected Path initWriter(Path root) throws IOException {
        Path workSheetPath = root.resolve(sheet.getName() + Const.Suffix.CSV);
        writer = CSVUtil.newWriter(workSheetPath);
        return workSheetPath;
    }

    /**
     * Write worksheet header data
     *
     * @throws IOException if I/O error occur
     */
    protected void writeBefore() throws IOException {
        // The header columns
        Sheet.Column[] columns = sheet.getHeaderColumns();
        boolean noneHeader = columns == null || columns.length == 0;

        if (!noneHeader) {
            for (Sheet.Column hc : columns) {
                writer.write(isNotEmpty(hc.getName()) ? hc.getName() : hc.key);
            }
            writer.newLine();
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
                case SST:
                    writer.write(cell.sv);
                    break;
                case NUMERIC:
                    writer.write(cell.nv);
                    break;
                case LONG:
                    writer.write(cell.lv);
                    break;
                case DOUBLE:
                    writer.write(cell.dv);
                    break;
                case BOOL:
                    writer.write(cell.bv);
                    break;
                case DECIMAL:
                    writer.write(cell.mv.toString());
                    break;
                case CHARACTER:
                    writer.writeChar(cell.cv);
                    break;
                case DATE:
                    writer.write(DateUtil.toDateString(DateUtil.toDate(cell.nv)));
                    break;
                case DATETIME:
                    writer.write(DateUtil.toString(DateUtil.toDate(cell.dv)));
                    break;
                case TIME:
                    writer.write(DateUtil.toDate(cell.dv).toString());
                    break;
                default:
                    writer.writeEmpty();
            }
        }
        writer.newLine();
    }
}

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

package org.ttzero.excel.entity.e7;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.ttzero.excel.entity.ICellValueAndStyle;
import org.ttzero.excel.entity.IDrawingsWriter;
import org.ttzero.excel.entity.Picture;
import org.ttzero.excel.entity.Watermark;
import org.ttzero.excel.entity.style.Border;
import org.ttzero.excel.entity.style.Fill;
import org.ttzero.excel.entity.style.Font;
import org.ttzero.excel.entity.style.PatternType;
import org.ttzero.excel.entity.style.Styles;
import org.ttzero.excel.manager.RelManager;
import org.ttzero.excel.manager.TopNS;
import org.ttzero.excel.entity.Column;
import org.ttzero.excel.entity.ExcelWriteException;
import org.ttzero.excel.entity.IWorksheetWriter;
import org.ttzero.excel.entity.Panes;
import org.ttzero.excel.entity.Relationship;
import org.ttzero.excel.entity.Row;
import org.ttzero.excel.entity.RowBlock;
import org.ttzero.excel.entity.SharedStrings;
import org.ttzero.excel.entity.Sheet;
import org.ttzero.excel.manager.Const;
import org.ttzero.excel.validation.Validation;
import org.ttzero.excel.reader.Cell;
import org.ttzero.excel.reader.Dimension;
import org.ttzero.excel.reader.Grid;
import org.ttzero.excel.reader.GridFactory;
import org.ttzero.excel.util.ExtBufferedWriter;
import org.ttzero.excel.util.FileSignatures;
import org.ttzero.excel.util.FileUtil;
import org.ttzero.excel.util.StringUtil;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.io.ObjectInputStream;
import java.io.ObjectOutputStream;
import java.io.OutputStream;
import java.io.OutputStreamWriter;
import java.lang.reflect.InvocationTargetException;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.net.HttpURLConnection;
import java.net.URL;
import java.nio.ByteBuffer;
import java.nio.ByteOrder;
import java.nio.channels.SeekableByteChannel;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.nio.file.StandardOpenOption;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.function.BiConsumer;
import java.util.stream.Collectors;

import static org.ttzero.excel.entity.Sheet.int2Col;
import static org.ttzero.excel.entity.Sheet.toCoordinate;
import static org.ttzero.excel.reader.Cell.BINARY;
import static org.ttzero.excel.reader.Cell.BOOL;
import static org.ttzero.excel.reader.Cell.BYTE_BUFFER;
import static org.ttzero.excel.reader.Cell.CHARACTER;
import static org.ttzero.excel.reader.Cell.DATE;
import static org.ttzero.excel.reader.Cell.DATETIME;
import static org.ttzero.excel.reader.Cell.DECIMAL;
import static org.ttzero.excel.reader.Cell.DOUBLE;
import static org.ttzero.excel.reader.Cell.FILE;
import static org.ttzero.excel.reader.Cell.INLINESTR;
import static org.ttzero.excel.reader.Cell.INPUT_STREAM;
import static org.ttzero.excel.reader.Cell.LONG;
import static org.ttzero.excel.reader.Cell.NUMERIC;
import static org.ttzero.excel.reader.Cell.REMOTE_URL;
import static org.ttzero.excel.reader.Cell.SST;
import static org.ttzero.excel.reader.Cell.TIME;
import static org.ttzero.excel.reader.Cell.UNALLOCATED;
import static org.ttzero.excel.util.ExtBufferedWriter.stringSize;
import static org.ttzero.excel.util.FileUtil.exists;
import static org.ttzero.excel.util.StringUtil.isNotEmpty;

/**
 * XML工作表输出
 *
 * @author guanquan.wang at 2019-04-22 16:31
 */
@TopNS(prefix = {"", "r"}, value = "worksheet"
    , uri = {Const.SCHEMA_MAIN, Const.Relationship.RELATIONSHIP})
public class XMLWorksheetWriter implements IWorksheetWriter {
    /**
     * LOGGER
     */
    protected final Logger LOGGER = LoggerFactory.getLogger(getClass());

    // the storage path
    protected Path workSheetPath, mediaPath;
    protected ExtBufferedWriter bw;
    protected Sheet sheet;
    protected Column[] columns;
    protected SharedStrings sst;
    /**
     * 全局样式，为工作表的一个指针
     */
    protected Styles styles;
    protected int startRow // The first data-row index
        , startHeaderRow // The first header row index
        , totalRows
        , sheetDataReady // Temporary code to increase compatibility with older versions
        ;
    /**
     * If there are any auto-width columns
     */
    protected boolean includeAutoWidth, ready;
    /**
     * Picture and Chart Support
     */
    protected IDrawingsWriter drawingsWriter;
    /**
     * A progress window
     */
    protected BiConsumer<Sheet, Integer> progressConsumer;

    // 自适应列宽专用
    protected double[] columnWidths;
    /**
     * 关系管理器（worksheet的副本）
     */
    protected RelManager relManager;
    /**
     * 超链接管理
     */
    protected Map<String, List<String>> hyperlinkMap;

    public XMLWorksheetWriter() { }

    public XMLWorksheetWriter(Sheet sheet) {
        this.sheet = sheet;
        this.relManager = sheet != null ? sheet.getRelManager() : null;
    }

    /**
     * Write a row block
     *
     * @param path the storage path
     * @throws IOException if I/O error occur
     */
    @Override
    public void writeTo(Path path) throws IOException {
        initWriter(path);

        // Get the first block
        RowBlock rowBlock = sheet.nextBlock();

        // write before
        writeBefore();

        // Write body data
        beforeSheetData(sheet.getNonHeader() == 1);

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

        totalRows = rowBlock.getTotal();
    }

    protected Path initWriter(Path root) throws IOException {
        this.workSheetPath = root.resolve("worksheets");
        if (!exists(this.workSheetPath)) {
            FileUtil.mkdir(workSheetPath);
        }

        Path sheetPath = workSheetPath.resolve(sheet.getFileName());
        // Already initialized
        if (ready) return sheetPath;

        this.bw = new ExtBufferedWriter(Files.newBufferedWriter(sheetPath, StandardCharsets.UTF_8));

        if (sst == null) this.sst = sheet.getWorkbook().getSharedStrings();
        if (styles == null) this.styles = sheet.getWorkbook().getStyles();

        // Check the first row index
        startHeaderRow = sheet.getStartRowNum();
        if (startHeaderRow <= 0)
            throw new IndexOutOfBoundsException("The start row index must be greater than 0, current = " + startHeaderRow);
        if (getRowLimit() <= startHeaderRow)
            throw new IndexOutOfBoundsException("The start row index must be less than row-limit, current(" + startHeaderRow + ") >= limit(" + getRowLimit() + ")");
        startRow = startHeaderRow;

        // Init progress window
        progressConsumer = sheet.getProgressConsumer();

        // Relationship
        if (relManager == null) relManager = sheet.getRelManager();

        if (hyperlinkMap == null) hyperlinkMap = new HashMap<>();

        ready = true;
        LOGGER.debug("{} WorksheetWriter initialization completed.", sheet.getName());
        return sheetPath;
    }

    /**
     * Rest worksheet
     *
     * @param sheet the worksheet
     * @return self
     */
    @Override
    public IWorksheetWriter setWorksheet(Sheet sheet) {
        this.sheet = sheet;
        this.relManager = sheet != null ? sheet.getRelManager() : null;
        return this;
    }

    @Override
    public IWorksheetWriter clone() {
        IWorksheetWriter copy;
        try {
            copy = (IWorksheetWriter) super.clone();
        } catch (CloneNotSupportedException e) {
            ObjectOutputStream oos = null;
            ObjectInputStream ois = null;
            try {
                ByteArrayOutputStream bos = new ByteArrayOutputStream();
                oos = new ObjectOutputStream(bos);
                oos.writeObject(this);

                ois = new ObjectInputStream(new ByteArrayInputStream(bos.toByteArray()));
                copy = (IWorksheetWriter) ois.readObject();
            } catch (IOException | ClassNotFoundException e1) {
                try {
                    copy = getClass().getConstructor(Sheet.class).newInstance(sheet);
                } catch (NoSuchMethodException | IllegalAccessException
                    | InstantiationException | InvocationTargetException e2) {
                    throw new ExcelWriteException(e2);
                }
            } finally {
                FileUtil.close(oos);
                FileUtil.close(ois);
            }
        }
        XMLWorksheetWriter e = (XMLWorksheetWriter) copy;
        e.sheetDataReady = 0;
        e.totalRows = 0;
        e.drawingsWriter = null;
        e.ready = false;
        e.bw = null;
        return copy;
    }

    /**
     * Returns the worksheet name
     *
     * @return name of worksheet
     */
    @Override
    public String getFileSuffix() {
        return Const.Suffix.XML;
    }

    /**
     * Override this method to modify the maximum number
     * of rows per page, this value contains the header
     * row and data rows
     * <p>
     * eq: limit is 100 means data has 99 rows
     *
     * @return the row limit
     */
    @Override
    public int getRowLimit() {
        return Const.Limit.MAX_ROWS_ON_SHEET;
    }

    /**
     * The Worksheet column limit
     *
     * @return the limit
     */
    @Override
    public int getColumnLimit() {
        return Const.Limit.MAX_COLUMNS_ON_SHEET;
    }

    /**
     * Write worksheet header data
     *
     * @throws IOException if I/O error occur
     */
    protected void writeBefore() throws IOException {
        // The header columns
        columns = sheet.getAndSortHeaderColumns();
        boolean nonHeader = sheet.getNonHeader() == 1;

        // 收集表头信息
        collectHeaderColumns();

        bw.write(Const.EXCEL_XML_DECLARATION);
        // Declaration
        bw.newLine();
        // Root node
        writeRootNode();

        // Dimension
        writeDimension();

        // SheetViews default value
        writeSheetViews();

        // Default row height and width
        int fillSpace = 17; // Column width xxx.xx (6byte) + hidden property (11byte)
        BigDecimal width = BigDecimal.valueOf(!nonHeader ? sheet.getDefaultWidth() : 8D).add(new BigDecimal("0.65"));
        // Overflow column width limit
        if (width.compareTo(new BigDecimal(Const.Limit.COLUMN_WIDTH)) > 0) {
            width = new BigDecimal(Const.Limit.COLUMN_WIDTH);
        }
        String defaultWidth = width.setScale(2, RoundingMode.HALF_UP).toString();

        // SheetFormatPr
        writeSheetFormat();

        // cols
        writeCols(fillSpace, defaultWidth);

//        // Initialization DrawingsWriter
//        initDrawingsWriter();
    }

    /**
     * Write the header row
     *
     * @return row number
     * @throws IOException if I/O error occur
     */
    protected int writeHeaderRow() throws IOException {
        // Write header
        int rowIndex = 0, subColumnSize = columns[0].subColumnSize();
        Column[][] columnsArray = new Column[columns.length][];
        for (int i = 0; i < columns.length; i++) {
            columnsArray[i] = columns[i].toArray();
        }
        // Merge cells if exists
        @SuppressWarnings("unchecked")
        List<Dimension> mergeCells = (List<Dimension>) sheet.getExtPropValue(Const.ExtendPropertyKey.MERGE_CELLS);
        Grid mergedGrid = mergeCells != null && !mergeCells.isEmpty() ? GridFactory.create(mergeCells) : null;
        Cell cell = new Cell();
        for (int i = subColumnSize - 1; i >= 0; i--) {
            // Custom row height
            double ht = getHeaderHeight(columnsArray, i);
            if (ht < 0) ht = sheet.getHeaderRowHeight();
            int row = startRow(rowIndex++, columns.length, ht);

            for (int j = 0, c = 0; j < columns.length; j++) {
                Column hc = columnsArray[j][i];
                cell.setString(isNotEmpty(hc.getName()) ? hc.getName() : mergedGrid != null && mergedGrid.test(i + 1, hc.getColNum()) && !isFirstMergedCell(mergeCells, i + 1, hc.getColNum()) ? null : hc.key);
                cell.xf = hc.getHeaderStyleIndex();
                writeString(cell, row, c++);
            }

            // Write header comments
            for (int j = 0; j < columns.length; j++) {
                Column hc = columnsArray[j][i];
                if (hc.headerComment != null) {
                    sheet.createComments().addComment(toCoordinate(row, hc.getColNum()), hc.headerComment);
                }
            }
            bw.write("</row>");
        }
        return subColumnSize;
    }

    /**
     * Write at after worksheet body
     *
     * @param total the total of rows
     * @throws IOException if I/O error occur
     */
    protected void writeAfter(int total) throws IOException {
        // End target --sheetData
        bw.write("</sheetData>");

        // 写完数据后调用工作表处理全局属性
        sheet.afterSheetDataWriter(total);

        // Auto Filter
        Dimension autoFilter = (Dimension) sheet.getExtPropValue(Const.ExtendPropertyKey.AUTO_FILTER);
        if (autoFilter != null) {
            bw.write("<autoFilter ref=\"");
            bw.write(autoFilter.toString());
            bw.write("\"/>");
        }

        // Merge cells
        writeMergeCells();

        // Others
        afterSheetData();

        // End target
        if (getClass().isAnnotationPresent(TopNS.class)) {
            TopNS topNS = getClass().getAnnotation(TopNS.class);
            bw.write("</");
            bw.write(topNS.value());
            bw.write('>');
        } else {
            bw.write("</worksheet>");
        }
        LOGGER.debug("Sheet [{}] writing completed, total rows: {}", sheet.getName(), total);
    }

    @Override
    public void writeData(RowBlock rowBlock) throws IOException {
        if (!ready) {
            Path path = sheet.getWorkbook().getWorkbookWriter().writeBefore();
            // Init
            initWriter(path.resolve("xl"));
            // write before
            writeBefore();
            // Write body data
            beforeSheetData(sheet.getNonHeader() == 1);
        }

        if (progressConsumer == null) writeRowBlock(rowBlock);
        else writeRowBlockFireProgress(rowBlock);

        totalRows = rowBlock.getTotal();
    }

    /**
     * 返回数据样式转换器
     *
     * @return 如果有斑马线则返回 {@link XMLZebraLineCellValueAndStyle}否则返回{@link XMLCellValueAndStyle}
     */
    @Override
    public ICellValueAndStyle getCellValueAndStyle() {
        int zebraFillStyle = sheet.getZebraFillStyle();
        return zebraFillStyle > 0 ? new XMLZebraLineCellValueAndStyle(zebraFillStyle) : new XMLCellValueAndStyle();
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
     * Write begin of row
     *
     * @param rows    the row index (zero base)
     * @param columns the column length
     * @param rowHeight the row height
     * @return the row index (one base)
     * @throws IOException if I/O error occur
     */
    protected int startRow(int rows, int columns, Double rowHeight) throws IOException {
        // Row number
        int r = rows + startRow;

        bw.write("<row r=\"");
        bw.writeInt(r);
        // default data row height 16.5
        if (rowHeight != null && rowHeight >= 0D) {
            bw.write("\" customHeight=\"1\" ht=\"");
            bw.write(rowHeight);
        }
        if (this.columns.length > 0) {
            bw.write("\" spans=\"");
            bw.writeInt(this.columns[0].getColNum());
            bw.write(':');
            bw.writeInt(this.columns[this.columns.length - 1].getColNum());
        } else {
            bw.write("\" spans=\"1:");
            bw.writeInt(columns);
        }
        bw.write("\">");
        return r;
    }

    /**
     * 写行的起始属性
     *
     * @param row 行对象{@link Row}
     * @return 行号
     * @throws IOException 出现输出异常
     */
    protected int startRow(Row row) throws IOException {
        // Row number
        int r = row.getIndex() + startRow;

        bw.write("<row r=\"");
        bw.writeInt(r);
        Double rowHeight = row.getHeight();
        // default data row height 16.5
        if (rowHeight != null && rowHeight >= 0D) {
            bw.write("\" customHeight=\"1\" ht=\"");
            bw.write(rowHeight);
        }
        if (row.lc - row.fc >= 1) {
            bw.write("\" spans=\"");
            bw.writeInt(row.fc + 1);
            bw.write(':');
            bw.writeInt(row.lc);
        }
        else if (this.columns.length > 0) {
            bw.write("\" spans=\"");
            bw.writeInt(this.columns[0].getColNum());
            bw.write(':');
            bw.writeInt(this.columns[this.columns.length - 1].getColNum());
        }
        // 隐藏行
        if (row.isHidden()) bw.write("\" hidden=\"1");
        // 层级
        if (row.getOutlineLevel() != null && row.getOutlineLevel().compareTo(0) > 0) {
            bw.write("\" outlineLevel=\"");
            bw.writeInt(row.getOutlineLevel());
        }
        return r;
    }

    /**
     * Write a row data
     *
     * @param row a row data
     * @throws IOException if I/O error occur
     */
    protected void writeRow(Row row) throws IOException {
        Cell[] cells = row.getCells();
        int r = startRow(row);
        if (row.lc > row.fc) {
            bw.write("\">");

            // 循环写单元格
            for (int i = row.fc; i < row.lc; i++) writeCell(cells[i], r, i);

            bw.write("</row>");
        } else bw.write("\"/>");
    }

    /**
     * 写单元格
     *
     * @param cell 单元格
     * @param row 行号
     * @param col 列下标，不等同于列号，列号通过{@link Column#getColNum()}获取
     * @throws IOException 出现输出异常
     */
    protected void writeCell(Cell cell, int row, int col) throws IOException {
        boolean valueOnly = cell.mediaType <= UNALLOCATED;
        // 写值
        switch (cell.t) {
            case INLINESTR:
            case SST:
                if (valueOnly) writeString(cell, row, col);
                else writeNull(cell, row, col);                  break;
            case NUMERIC:
            case LONG:
            case DATE:
            case DATETIME:
            case DOUBLE:
            case TIME:
            case DECIMAL:   writeNumeric(cell, row, col);        break;
            case BOOL:      writeBool(cell, row, col);           break;
            case CHARACTER: writeChar(cell, row, col);           break;
            default:        writeNull(cell, row, col);
        }

        // 图片
        if (!valueOnly) {
            switch (cell.mediaType) {
                case REMOTE_URL  : writeRemoteMedia(cell.stringVal, row, col); break;
                case FILE        : writeFile(cell.path, row, col);             break;
                case INPUT_STREAM: writeStream(cell.isv, row, col);            break;
                case BINARY      : writeBinary(cell.binary, row, col);         break;
                case BYTE_BUFFER : writeBinary(cell.byteBuffer, row, col);     break;
            }
        }
    }

    /**
     * 写字符串
     *
     * @param cell   单元格信息
     * @param row    行号
     * @param col    列下标
     * @throws IOException if I/O error occur
     */
    protected void writeString(Cell cell, int row, int col) throws IOException {
        Column hc = getColumn(col);
        bw.write("<c r=\"");
        bw.write(int2Col(hc.getColNum()));
        bw.writeInt(row);

        String s = cell.stringVal;
        boolean notEmpty = s != null && !s.isEmpty();

        // 超链接
        if (cell.h && notEmpty) {
            Relationship rel = relManager.add(new Relationship(s, Const.Relationship.HYPERLINK).setTargetMode("External"));
            List<String> dim = hyperlinkMap.computeIfAbsent(rel.getId(), k -> new ArrayList<>());
            dim.add(toCoordinate(row, col + 1));
        }

        if (cell.xf > 0) {
            bw.write("\" s=\"");
            bw.writeInt(cell.xf);
        }

        if (cell.f) {
            bw.write("\" t=\"str\"><f>");
            bw.escapeWrite(cell.formula);
            bw.write("</f>");
            if (notEmpty) {
                bw.write("<v>");
                bw.escapeWrite(s);
                bw.write("</v>");
            }
            bw.write("</c>");
        } else if (notEmpty) {
            int i;
            if (hc.isShare() && (i = sst.get(s)) >= 0) {
                bw.write("\" t=\"s\"><v>");
                bw.writeInt(i);
                bw.write("</v></c>");
            } else {
                bw.write("\" t=\"inlineStr\"><is><t>");
                bw.escapeWrite(s); // escape text
                bw.write("</t></is></c>");
            }
        } else bw.write("\"/>");

        // TODO optimize If auto-width
        if (hc.getAutoSize() == 1) {
            double ln;
            if (columnWidths[col] < (ln = stringWidth(s, cell.xf))) columnWidths[col] = ln;
        }
    }

    /**
     * 写数字
     *
     * @param cell   单元格信息
     * @param row    行号
     * @param col    列下标
     * @throws IOException if I/O error occur
     */
    protected void writeNumeric(Cell cell, int row, int col) throws IOException {
        Column hc = getColumn(col);
        bw.write("<c r=\"");
        bw.write(int2Col(hc.getColNum()));
        bw.writeInt(row);
        if (cell.xf > 0) {
            bw.write("\" s=\"");
            bw.writeInt(cell.xf);
        }
        bw.write("\">");
        if (cell.f) {
            bw.write("<f>");
            bw.escapeWrite(cell.formula);
            bw.write("</f>");
        }
        bw.write("<v>");
        boolean autoSize = hc.getAutoSize() == 1;
        String s = null;
        switch (cell.t) {
            case NUMERIC:
                bw.writeInt(cell.intVal);
                if (autoSize) s = Integer.toString(cell.intVal);
                break;
            case LONG:
                bw.write(cell.longVal);
                if (autoSize) s = Long.toString(cell.longVal);
                break;
            case DATE:
            case DATETIME:
            case DOUBLE:
            case TIME:
                bw.write(cell.doubleVal);
                if (autoSize) s = Double.toString(cell.doubleVal);
                break;
            case DECIMAL:
                bw.write(s = cell.decimal.toString());
                break;
        }
        bw.write("</v></c>");

        if (autoSize && s != null) {
            double n;
            if (hc.getNumFmt() != null) {
                if (columnWidths[col] < (n = hc.getNumFmt().calcNumWidth(s.length(), getFont(cell.xf)))) columnWidths[col] = n;
            }
            else if (columnWidths[col] < (n = stringWidth(s, cell.xf))) columnWidths[col] = n;
        }
    }

    /**
     * 写布尔值
     *
     * @param cell   单元格信息
     * @param row    行号
     * @param col    列下标
     * @throws IOException if I/O error occur
     */
    protected void writeBool(Cell cell, int row, int col) throws IOException {
        Column hc = getColumn(col);
        bw.write("<c r=\"");
        bw.write(int2Col(hc.getColNum()));
        bw.writeInt(row);
        bw.write("\" t=\"b");
        if (cell.xf > 0) {
            bw.write("\" s=\"");
            bw.writeInt(cell.xf);
        }
        bw.write("\">");
        if (cell.f) {
            bw.write("<f>");
            bw.escapeWrite(cell.formula);
            bw.write("</f>");
        }
        bw.write("<v>");
        bw.writeInt(cell.boolVal ? 1 : 0);
        bw.write("</v></c>");

        // TODO optimize If auto-width
        if (hc.getAutoSize() == 1) {
            double ln;
            if (columnWidths[col] < (ln = stringWidth(Boolean.toString(cell.boolVal), cell.xf))) columnWidths[col] = ln;
        }
    }

    /**
     * 写字符
     *
     * @param cell   单元格信息
     * @param row    行号
     * @param col    列下标
     * @throws IOException if I/O error occur
     */
    protected void writeChar(Cell cell, int row, int col) throws IOException {
        Column hc = getColumn(col);
        bw.write("<c r=\"");
        bw.write(int2Col(hc.getColNum()));
        bw.writeInt(row);
        if (cell.xf > 0) {
            bw.write("\" s=\"");
            bw.writeInt(cell.xf);
        }
        char c = cell.charVal;
        if (cell.f) {
            bw.write("\" t\"str\"><f>");
            bw.escapeWrite(cell.formula);
            bw.write("</f><v>");
            bw.escapeWrite(c);
            bw.write("</v></c>");
        } else if (hc.isShare()) {
            bw.write("\" t=\"s\"><v>");
            bw.writeInt(sst.get(c));
            bw.write("</v></c>");
        } else {
            bw.write("\" t=\"inlineStr\"><is><t>");
            bw.escapeWrite(c);
            bw.write("</t></is></c>");
        }
        // TODO optimize If auto-width
        if (hc.getAutoSize() == 1) {
            Font font = getFont(cell.xf);
            double n = (c > 0x4E00 ? font.getSize() : font.getFontMetrics().charWidth(c)) / 6.0D * 1.02D;
            if (columnWidths[col] < n) columnWidths[col] = n;
        }
    }

    /**
     * 写空值
     *
     * @param cell   单元格信息
     * @param row    行号
     * @param col    列下标
     * @throws IOException if I/O error occur
     */
    protected void writeNull(Cell cell, int row, int col) throws IOException {
        // 有公式、边框或填充时写空单元格
        if (cell.xf == 0 && !cell.f) return;
        int style = styles.getStyleByIndex(cell.xf);
        Fill fill = styles.getFill(style);
        Border border = styles.getBorder(style);
        if (fill != null && fill.getPatternType() != PatternType.none || border != null && border.isEffectiveBorder() || cell.f) {
            bw.write("<c r=\"");
            bw.write(int2Col(getColumn(col).getColNum()));
            bw.writeInt(row);
            bw.write("\" s=\"");
            bw.writeInt(cell.xf);
            if (cell.f) {
                bw.write("\"><f>");
                bw.escapeWrite(cell.formula);
                bw.write("</f></c>");
            } else bw.write("\"/>");
        }
    }

    /**
     * Write binary file
     *
     * @param bytes  the binary data
     * @param row    the row index
     * @param column the column index
     * @throws IOException if I/O error occur
     */
    protected void writeBinary(byte[] bytes, int row, int column) throws IOException {
        // Test file signatures
        FileSignatures.Signature signature = FileSignatures.test(ByteBuffer.wrap(bytes));
        if (signature == null || !signature.isTrusted()) {
            LOGGER.warn("File types that are not allowed");
            return;
        }
        // 实例化drawingsWriter
        if (drawingsWriter == null) createDrawingsWriter();
        int id = sheet.getWorkbook().incrementMediaCounter();
        String name = "image" + id + "." + signature.extension;
        // Store in disk
        Files.write(mediaPath.resolve(name), bytes, StandardOpenOption.CREATE_NEW);

        // Write picture
        writePictureDirect(id, name, column, row, signature);
    }

    /**
     * Write binary file
     *
     * @param byteBuffer  the binary data
     * @param row    the row index
     * @param column the column index
     * @throws IOException if I/O error occur
     */
    protected void writeBinary(ByteBuffer byteBuffer, int row, int column) throws IOException {
        int position = byteBuffer.position();
        // Test file signatures
        FileSignatures.Signature signature = FileSignatures.test(byteBuffer);
        if (signature == null || !signature.isTrusted()) {
            LOGGER.warn("File types that are not allowed");
            return;
        }
        // 实例化drawingsWriter
        if (drawingsWriter == null) createDrawingsWriter();
        int id = sheet.getWorkbook().incrementMediaCounter();
        String name = "image" + id + "." + signature.extension;
        // Reset buffer position
        byteBuffer.position(position);
        // Store in disk
        SeekableByteChannel channel = Files.newByteChannel(mediaPath.resolve(name), StandardOpenOption.WRITE, StandardOpenOption.CREATE_NEW);
        channel.write(byteBuffer);
        channel.close();

        // Write picture
        writePictureDirect(id, name, column, row, signature);
    }

    /**
     * Write file value
     *
     * @param path   the picture file
     * @param row    the row index
     * @param column the column index
     * @throws IOException if I/O error occur
     */
    protected void writeFile(Path path, int row, int column) throws IOException {
        // Test file signatures
        FileSignatures.Signature signature = FileSignatures.test(path);
        if (!signature.isTrusted()) {
            LOGGER.warn("File types that are not allowed");
            return;
        }
        // 实例化drawingsWriter
        if (drawingsWriter == null) createDrawingsWriter();
        int id = sheet.getWorkbook().incrementMediaCounter();
        String name = "image" + id + "." + signature.extension;
        // Store
        Files.copy(path, mediaPath.resolve(name), StandardCopyOption.REPLACE_EXISTING);

        // Write picture
        writePictureDirect(id, name, column, row, signature);
    }

    /**
     * Write stream value
     *
     * @param stream  the picture input-stream
     * @param row    the row index
     * @param column the column index
     * @throws IOException if I/O error occur
     */
    protected void writeStream(InputStream stream, int row, int column) throws IOException {
        byte[] bytes = new byte[1 << 13];
        int n;

        OutputStream os = null;
        try {
            n = stream.read(bytes);
            // Empty stream
            if (n <= 0) return;
            FileSignatures.Signature signature = FileSignatures.test(ByteBuffer.wrap(bytes, 0, n));
            if (signature == null || !signature.isTrusted()) {
                LOGGER.warn("File types that are not allowed");
                return;
            }
            // 实例化drawingsWriter
            if (drawingsWriter == null) createDrawingsWriter();
            int id = sheet.getWorkbook().incrementMediaCounter();
            String name = "image" + id + "." + signature.extension;
            os = Files.newOutputStream(mediaPath.resolve(name));
            os.write(bytes, 0, n);

            if (n == bytes.length) {
                while ((n = stream.read(bytes)) > 0)
                    os.write(bytes, 0, n);
            }

            // Write picture
            writePictureDirect(id, name, column, row, signature);
        } finally {
            try {
                stream.close();
            } catch (IOException e) { } // Ignore
            if (os != null) {
                try {
                    os.close();
                } catch (IOException e) { } // Ignore
            }
        }
    }

    /**
     * Write remote media value
     *
     * @param url  remote url
     * @param row    the row index
     * @param column the column index
     * @throws IOException if I/O error occur
     */
    protected void writeRemoteMedia(String url, int row, int column) throws IOException {
        Picture picture = createPicture(column, row);
        picture.id = sheet.getWorkbook().incrementMediaCounter();

        // 实例化drawingsWriter
        if (drawingsWriter == null) createDrawingsWriter();
        // Async Drawing
        drawingsWriter.asyncDrawing(picture);

        // Supports asynchronous download
        downloadRemoteResource(picture, url);
    }

    /**
     * 写{@link Picture}
     *
     * @param picture 图像对象
     * @throws IOException 输出异常
     */
    @Override
    public void writePicture(Picture picture) throws IOException {
        if (picture.localPath == null) return;
        // Test file signatures
        FileSignatures.Signature signature = FileSignatures.test(picture.localPath);
        if (!signature.isTrusted()) {
            LOGGER.warn("File types that are not allowed");
            return;
        }
        // 实例化drawingsWriter
        if (drawingsWriter == null) createDrawingsWriter();
        int id = sheet.getWorkbook().incrementMediaCounter();
        picture.id = id;
        String name = "image" + id + "." + signature.extension;
        picture.picName = name;
        picture.size = signature.width << 16 | signature.height;
        // Store
        Files.copy(picture.localPath, mediaPath.resolve(name), StandardCopyOption.REPLACE_EXISTING);

        // Write picture
        // Drawing
        drawingsWriter.drawing(picture);

        // Add global contentType
        sheet.getWorkbook().addContentType(new ContentType.Default(signature.contentType, signature.extension));
    }

    /**
     * Resize column width
     *
     * @param path the sheet temp path
     * @param rows total of rows
     * @throws IOException if I/O error occur
     */
    protected void resizeColumnWidth(File path, int rows) throws IOException {
        if (includeAutoWidth) {
            // There has no column to reset width
            if (columns.length <= 0 || rows <= 0) return;
            // Collect column width
            for (int i = 0; i < columns.length; i++) {
                Column hc = columns[i];
                int k = hc.getAutoSize();
                // If fixed width or media cell
                if (k == 2 || hc.getColumnType() == 1) {
                    double width = hc.width >= 0.0D ? hc.width : sheet.getDefaultWidth();
                    hc.width = BigDecimal.valueOf(Math.min(width + 0.65D, Const.Limit.COLUMN_WIDTH)).setScale(2, RoundingMode.HALF_UP).doubleValue();
                    continue;
                }
                double len = columnWidths[i] > 0 ? columnWidths[i] : sheet.getDefaultWidth();
                double width = (sheet.getNonHeader() == 1 ? len : Math.max(stringWidth(hc.name, hc.getHeaderStyleIndex() == -1 ? sheet.defaultHeadStyleIndex() : hc.getHeaderStyleIndex()), len)) + 1.86D;
                if (hc.width > 0.000001D) width = Math.min(width, hc.width + 0.65D);
                if (width > Const.Limit.COLUMN_WIDTH) width = Const.Limit.COLUMN_WIDTH;
                hc.width = BigDecimal.valueOf(width).setScale(2, RoundingMode.HALF_UP).doubleValue();
            }
        }

        if (bw != null) {
            try {
                bw.close();
            } catch (IOException ex) {
                // Ignore
            }
        }

        XMLWorksheetWriter _writer = new XMLWorksheetWriter(sheet);
        _writer.totalRows = totalRows;
        _writer.startRow = startRow;
        _writer.startHeaderRow = startHeaderRow;
        _writer.includeAutoWidth = includeAutoWidth;
        _writer.styles = styles;
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        _writer.bw = new ExtBufferedWriter(new OutputStreamWriter(baos, StandardCharsets.UTF_8));
        _writer.writeBefore();
        _writer.bw.close();

        // Resize each column width ...
        try (SeekableByteChannel channel = Files.newByteChannel(path.toPath(), StandardOpenOption.WRITE, StandardOpenOption.READ)) {
            ByteBuffer buffer = ByteBuffer.wrap(baos.toByteArray());
            buffer.order(ByteOrder.LITTLE_ENDIAN);
            channel.write(buffer);
        }
    }

    /**
     * Release resources
     */
    @Override
    public void close() throws IOException {
        if (bw != null) {
            // write end
            writeAfter(totalRows);

            // Resize if include auto-width column
            resizeColumnWidth(workSheetPath.resolve(sheet.getFileName()).toFile(), totalRows);

            FileUtil.close(bw);
            bw = null;
        }
        // Write some final info
        sheet.afterSheetAccess(workSheetPath);

        // Close drawing writer
        FileUtil.close(drawingsWriter);
    }

    /**
     * Write the &lt;worksheet&gt; node
     *
     * @throws IOException if I/O error occur.
     */
    protected void writeRootNode() throws IOException {
        if (getClass().isAnnotationPresent(TopNS.class)) {
            TopNS topNS = getClass().getAnnotation(TopNS.class);
            bw.write('<');
            bw.write(topNS.value());
            String[] prefixs = topNS.prefix(), urls = topNS.uri();
            for (int i = 0, len = prefixs.length; i < len; ) {
                bw.write(" xmlns");
                if (prefixs[i] != null && !prefixs[i].isEmpty()) {
                    bw.write(':');
                    bw.write(prefixs[i]);
                }
                bw.write("=\"");
                bw.write(urls[i]);
                if (++i < len) {
                    bw.write('"');
                }
            }
        } else {
            bw.write("<worksheet xmlns=\"");
            bw.write(Const.SCHEMA_MAIN);
        }
        bw.write("\">");
    }

    /**
     * Write the dimension of sheet, default value is {@code A1}
     *
     * @throws IOException if I/O error occur.
     */
    protected void writeDimension() throws IOException {
        bw.append("<dimension ref=\"");
        if (columns.length > 0) bw.write(int2Col(columns[0].getColNum()));
        else bw.write('A');
        bw.writeInt(startHeaderRow);
        int n = 11, size = totalRows; // fill 11 space
        if (size > 0 && columns.length > 0) {
            bw.write(':');
            n--;
            Column hc = columns[columns.length - 1];
            char[] col = int2Col(hc.getColNum());
            bw.write(col);
            n -= col.length;
            size = size + startRow - 1;
            bw.writeInt(size > getRowLimit() ? getRowLimit() : size);
            n -= stringSize(size);
        }
        bw.write('"');
        for (; n-->0;) bw.write(32); // Fill space
        bw.write("/>");
    }

    /**
     * Write the sheet views such as FreezeEnum, Default selection cell.
     *
     * @throws IOException if I/O error occur.
     */
    protected void writeSheetViews() throws IOException {
        bw.write("<sheetViews>");

        bw.write("<sheetView workbookViewId=\"0\"");
        // Default show grid lines
        if (!sheet.isShowGridLines()) bw.write(" showGridLines=\"0\"");
        Object o = sheet.getExtPropValue(Const.ExtendPropertyKey.ZOOM_SCALE);
        if (o instanceof Integer) {
            int scale = (Integer) o;
            bw.write(" zoomScale=\"");
            // Scale value between 10% to 400%
            bw.writeInt(scale < 10 ? 10 : Math.min(scale, 400));
            bw.write("\"");
        }
        // Default select the first worksheet
        if (sheet.getId() == 1) bw.write(" tabSelected=\"1\"");

        // Freeze Panes
        o = sheet.getExtPropValue(Const.ExtendPropertyKey.FREEZE);
        if (o instanceof Panes) {
            Panes freezePanes = (Panes) o;

            // Validity check
            if (freezePanes.row < 0 || freezePanes.col < 0) {
                throw new IllegalArgumentException("Negative number occur in freeze panes settings.");
            }

            if ((freezePanes.col | freezePanes.row) == 0) {
                bw.write("/>"); // Empty tag
            } else {
                bw.write(">");

                Dimension dim = new Dimension(freezePanes.row + 1, (short) (freezePanes.col + 1));
                // Freeze top row
                if (freezePanes.col == 0) {
                    bw.write("<pane ySplit=\"" + freezePanes.row + "\" topLeftCell=\"" + dim + "\" activePane=\"bottomLeft\" state=\"frozen\"/>");
                    bw.write("<selection pane=\"bottomLeft\" activeCell=\"" + dim + "\" sqref=\"" + dim + "\"/>");
                }
                // Freeze first column
                else if (freezePanes.row == 0) {
                    bw.write("<pane xSplit=\"" + freezePanes.col + "\" topLeftCell=\"" + dim + "\" activePane=\"topRight\" state=\"frozen\"/>");
                    bw.write("<selection pane=\"topRight\" activeCell=\"" + dim + "\" sqref=\"" + dim + "\"/>");
                }
                // Freeze panes
                else {
                    bw.write("<pane xSplit=\"" + freezePanes.col + "\" ySplit=\"" + freezePanes.row + "\" topLeftCell=\"" + dim + "\" activePane=\"bottomRight\" state=\"frozen\"/>");
                    bw.write("<selection pane=\"topRight\" activeCell=\"" + new Dimension(1, dim.firstColumn) + "\" sqref=\"" + new Dimension(1, dim.firstColumn) + "\"/>");
                    bw.write("<selection pane=\"bottomLeft\" activeCell=\"" + new Dimension(dim.firstRow, (short) 1) + "\" sqref=\"" + new Dimension(dim.firstRow, (short) 1) + "\"/>");
                    bw.write("<selection pane=\"bottomRight\" activeCell=\"" + dim + "\" sqref=\"" + dim + "\"/>");
                }
                bw.write("</sheetView>");
            }
        }
        // Move the head row to the top
        else if (sheet.isScrollToVisibleArea() && startHeaderRow > 1) {
            bw.write(" topLeftCell=\"");
            char[] cols = int2Col(columns[0].getColNum());
            bw.write(cols);
            bw.writeInt(startHeaderRow);
            bw.write("\">");
            bw.write("<selection activeCell=\"");
            bw.write(cols);
            bw.writeInt(startHeaderRow);
            bw.write("\" sqref=\"");
            bw.write(cols);
            bw.writeInt(startHeaderRow);
            bw.write("\"/></sheetView>");
        }
        else {
            bw.write("/>"); // Empty tag
        }
        bw.write("</sheetViews>");
    }

    /**
     * Write the sheet format
     *
     * @throws IOException if I/O error occur.
     */
    protected void writeSheetFormat() throws IOException {
        int n = 0;
        BigDecimal defaultColWidth = null, defaultRowHeight = null;
        try {
            Object o;
            if ((o = sheet.getExtPropValue("defaultColWidth")) != null) {
                defaultColWidth = new BigDecimal(o.toString());
                n |= 1;
            }
            if ((o = sheet.getExtPropValue("defaultRowHeight")) != null) {
                defaultRowHeight = new BigDecimal(o.toString());
                n |= 2;
            }
        } catch (NumberFormatException e) {
            // Ignore
        }
        if (n > 0) {
            bw.write("<sheetFormatPr");
            if ((n & 1) == 1) {
                bw.write(" defaultColWidth=\"");
                bw.writeInt(defaultColWidth.intValue());
                bw.write("\"");
            }
            if ((n & 2) == 2) {
                bw.write(" defaultRowHeight=\"");
                bw.writeInt(defaultRowHeight.intValue());
                bw.write("\"");
            }
            bw.write("/>");
        }
    }

    /**
     * Write the default column info, The specified column width will be overwritten in these method.
     *
     * @param fillSpace The number of characters to pad when recalculating the width.
     * @param defaultWidth The default cell width, {@code 8.38} will be use if it not be setting.
     * @throws IOException if I/O error occur.
     */
    protected void writeCols(int fillSpace, String defaultWidth) throws IOException {
        if (columns.length > 0) {
            bw.write("<cols>");
            Column fCol = columns[0];
            String fWidth = fCol.width >= 0.0000001D ? new BigDecimal(fCol.width).setScale(2, RoundingMode.HALF_UP).toString() : defaultWidth;
            // 多个col时将相同属性的col进行压缩
            if (columns.length > 1) {
                for (int i = 1; i < columns.length; i++) {
                    Column col = columns[i], pCol = columns[i - 1];
                    String width = col.width >= 0.0000001D ? new BigDecimal(col.width).setScale(2, RoundingMode.HALF_UP).toString() : defaultWidth;
                    boolean lastColumn = i == columns.length - 1;
                    if (fCol.getAutoSize() == 1 || col.getAutoSize() == 1 || !width.equals(fWidth) || col.isHide() != fCol.isHide() || col.getColNum() - pCol.getColNum() > 1 || fCol.globalStyleIndex != col.globalStyleIndex) {
                        writeCol(fWidth, fCol.getColNum(), pCol.getColNum(), fillSpace, fCol.globalStyleIndex, fCol.isHide());
                        fWidth = width;
                        fCol = col;
                    }
                    if (lastColumn) writeCol(width, fCol.getColNum(), col.getColNum(), fillSpace, fCol.globalStyleIndex, col.isHide());
                }
            } else writeCol(fWidth, fCol.getColNum(), fCol.getColNum(), fillSpace, fCol.globalStyleIndex, fCol.isHide());
            bw.write("</cols>");
        }
    }

    protected void writeCol(String width, int min, int max, int fillSpace, int xf, boolean isHide) throws IOException {
        bw.write("<col customWidth=\"1\" width=\"");
        bw.write(width);
        int w = width.length();
        if (isHide) {
            bw.write("\" hidden=\"1");
            w += 11;
        }
        bw.write('"');
        for (int j = fillSpace - w; j-- > 0; ) bw.write(32); // Fill space
        bw.write(" min=\"");
        bw.writeInt(min);
        bw.write("\" max=\"");
        bw.writeInt(max);
        if (xf > 0) {
            bw.write("\" style=\"");
            bw.writeInt(xf);
        }
        bw.write("\" bestFit=\"1\"/>");
    }

    /**
     * Begin to write the Sheet data and the header row.
     *
     * @param nonHeader mark none header
     * @throws IOException if I/O error occur.
     */
    protected void beforeSheetData(boolean nonHeader) throws IOException {
        if (sheetDataReady > 0) return;
        // Start to write sheet data
        bw.write("<sheetData>");

        int headerRow = 0;
        // Write header rows
        if (!nonHeader && columns.length > 0) {
            headerRow = writeHeaderRow();
        }
        startRow = startHeaderRow + headerRow;
        sheetDataReady = 1;
    }

    /**
     * Append others customize info
     *
     * @throws IOException if I/O error occur.
     */
    protected void afterSheetData() throws IOException {
        // 数据验证
        @SuppressWarnings("unchecked")
        List<Validation> validations = (List<Validation>) sheet.getExtPropValue(Const.ExtendPropertyKey.DATA_VALIDATION);
        List<Validation> extList = null;
        if (validations != null && !validations.isEmpty()) {
            // Extension validation
            extList = validations.stream().filter(Validation::isExtension).collect(Collectors.toList());
            if (extList.size() < validations.size()) {
                bw.write("<dataValidations count=\"");
                bw.writeInt(validations.size() - extList.size());
                bw.write("\">");
                for (Validation e : validations) {
                    if (!e.isExtension()) bw.write(e.toString());
                }
                bw.write("</dataValidations>");
            }
        }

        // 超链接
        if (!hyperlinkMap.isEmpty()) {
            bw.write("<hyperlinks>");
            for (Map.Entry<String, List<String>> entry : hyperlinkMap.entrySet()) {
                for (String dim : entry.getValue()) {
                    bw.write("<hyperlink ref=\"");
                    bw.write(dim);
                    bw.write("\" r:id=\"");
                    bw.write(entry.getKey());
                    bw.write("\"/>");
                }
            }
            bw.write("</hyperlinks>");
        }

        Relationship r;
        // Drawings
        if (drawingsWriter != null) {
            r = relManager.getByType(Const.Relationship.DRAWINGS);
            if (r != null) {
                bw.write("<drawing r:id=\"");
                bw.write(r.getId());
                bw.write("\"/>");
            }
        }

        // vmlDrawing
        r = sheet.findRel("vmlDrawing");
        if (r != null) {
            bw.write("<legacyDrawing r:id=\"");
            bw.write(r.getId());
            bw.write("\"/>");
        }

        // Compatible processing
        else if (sheet.getComments() != null) {
            sheet.addRel(r = new Relationship("../drawings/vmlDrawing" + sheet.getId() + Const.Suffix.VML, Const.Relationship.VMLDRAWING));
            sheet.addRel(new Relationship("../comments" + sheet.getId() + Const.Suffix.XML, Const.Relationship.COMMENTS));

            bw.write("<legacyDrawing r:id=\"");
            bw.write(r.getId());
            bw.write("\"/>");
        }

        // 背景图片
        writeWatermark();

        // 扩展节点
        writeExtList(extList);
    }

    /**
     * 添加水印
     *
     * @throws IOException 无权限或磁盘空间不足
     */
    protected void writeWatermark() throws IOException {
        Watermark watermark = sheet.getWatermark();
        if (watermark == null || !watermark.canWrite()) {
            watermark = sheet.getWorkbook().getWatermark();
            sheet.setWatermark(watermark);
        }
        if (watermark != null && watermark.canWrite()) {
            Path media = workSheetPath.getParent().resolve("media");
            if (!exists(media)) Files.createDirectory(media);
            Path image = media.resolve("image" + sheet.getWorkbook().incrementMediaCounter() + watermark.getSuffix());

            Files.copy(watermark.get(), image);
            Relationship r = new Relationship("../media/" + image.getFileName(), Const.Relationship.IMAGE);
            sheet.addRel(r);

            bw.write("<picture r:id=\"");
            bw.write(r.getId());
            bw.write("\"/>");
        }
    }

    /**
     * 添加扩展节点
     *
     * @param extList 数据校验扩展节点
     * @throws IOException if I/O error occur.
     */
    protected void writeExtList(List<Validation> extList) throws IOException {
        // 扩展节点-当前只支持数据校验
        if (extList == null || extList.isEmpty()) return;
        bw.write("<extLst><ext xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\" uri=\"{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}\">");
        bw.write("<x14:dataValidations xmlns:xm=\"http://schemas.microsoft.com/office/excel/2006/main\" count=\"");
        bw.writeInt(extList.size());
        bw.write("\">");
        for (Validation e : extList) {
            bw.write(e.toString());
        }
        bw.write("</x14:dataValidations></ext></extLst>");
    }

    /**
     * Append merged cells if exists
     *
     * @throws IOException if I/O error occur.
     */
    protected void writeMergeCells() throws IOException {
        // Merge cells if exists
        @SuppressWarnings("unchecked")
        List<Dimension> mergeCells = (List<Dimension>) sheet.getExtPropValue(Const.ExtendPropertyKey.MERGE_CELLS);
        if (mergeCells != null && !mergeCells.isEmpty()) {
            bw.write("<mergeCells count=\"");
            bw.writeInt(mergeCells.size());
            bw.write("\">");
            for (Dimension dim : mergeCells) {
                bw.write("<mergeCell ref=\"");
                bw.write(dim.toString());
                bw.write("\"/>");
            }
            bw.write("</mergeCells>");
        }
    }

    /**
     * 计算文本在单元格的宽度，参考{@link sun.swing.SwingUtilities2#stringWidth}
     *
     * <p>Java的{@link java.awt.FontMetrics}计算中文字符有个问题，对于绝大多数英文字体计算出来的中文字符宽度都不准，
     * 在英文字体中显示的中文字体可能会默认显示"宋体"，默认字体与地区和操作系统相关这里只取简体中文的临近值约为字体大小，
     * 英文字符计算比较复杂每种字体显示的宽度差异很大，有较窄的字符{@code 'i','l',':'}也有较宽的字符{@code 'X','E','G'，’%'，‘@’}，
     * 对于英文字符统一使用{@code FontMetrics}计算。</p>
     *
     * <p>对于自动折行且自适应列宽的单元格则分别计算每一段文本宽度取最大值，这也是为什么不直接调用{@link java.awt.FontMetrics#stringWidth}
     * 的原因因为它并不会分段计算</p>
     *
     * <p>本方法计算的宽度在某些字体下计算出来的结果与实际显示效果可能有很大偏差，此时可以覆写本方法并进行特殊计算</p>
     *
     * @param s      文本
     * @param xf     单元格样式索引
     * @return 文本在excel的宽度
     */
    protected double stringWidth(String s, int xf) {
        if (StringUtil.isEmpty(s)) return 0.0D;
        Font font = getFont(xf);
        int fs = font.getSize(), w = 0;
        java.awt.FontMetrics fm = font.getFontMetrics();
        int len = s.length(), i = 0;
        char c;
        for (; i < len && w < 1500 && (c = s.charAt(i++)) != '\n'; w += c > 0x4E00 ? fs : fm.charWidth(c));
        // 如果包含回车则特殊处理
        if (i < len && w < 1500 && s.charAt(i - 1) == '\n') {
            int style = styles.getStyleByIndex(xf);
            // “自动折行”时计算每段长度取最大值
            if (Styles.hasWrapText(style)) {
                do {
                    int sectionWidth = 0;
                    for (; i < len && sectionWidth < 1500 && (c = s.charAt(i++)) != '\n'; sectionWidth += c > 0x4E00 ? fs : fm.charWidth(c));
                    if (sectionWidth > w) w = sectionWidth;
                } while (i < len && w < 1500);
            }
            // 非“自动折行”将显示为一行，宽度直接相加
            else for (; i < len && w < 1500; w += (c = s.charAt(i++)) > 0x4E00 ? fs : fm.charWidth(c));
        }
        return w / 6.0D * 1.02D;
    }

    /**
     * Test whether the coordinate is the first cell of the merged cell,
     * and use {@link Grid#test} to determine yes before calling this method
     *
     * @param mergeCells all merged cells
     * @param row the cell row
     * @param col the cell column
     * @return true if the is first cell in merged
     */
    public static boolean isFirstMergedCell(List<Dimension> mergeCells, int row, int col) {
        for (Dimension dim : mergeCells) {
            if (dim.checkRange(row, col) && row == dim.firstRow && col == dim.firstColumn) return true;
        }
        return false;
    }

    /**
     * Returns the maximum cell height
     *
     * @param columnsArray the header column array
     * @param row actual rows in Excel
     * @return cell height or -1
     */
    public double getHeaderHeight(Column[][] columnsArray, int row) {
        double h = -1D;
        for (Column[] cols : columnsArray) h = Math.max(cols[row].headerHeight, h);
        return h;
    }

//    /**
//     * Check if images are output
//     *
//     * @return true if any column type equals 1
//     */
//    protected boolean hasMedia() {
//        boolean hasMedia = false;
//        for (Column column : columns) {
//            hasMedia = column.getColumnType() == 1;
//            if (hasMedia) break;
//        }
//        return hasMedia;
//    }
//
//    /**
//     * Initialization DrawingsWriter
//     *
//     * @throws IOException if I/O error occur
//     */
//    protected void initDrawingsWriter() throws IOException {
//        if (hasMedia()) {
//            createDrawingsWriter();
//        }
//    }

    /**
     * Create drawing writer and add relationship
     *
     * @return {@link IDrawingsWriter}
     * @throws IOException if I/O error occur.
     */
    protected IDrawingsWriter createDrawingsWriter() throws IOException {
        if (mediaPath == null) mediaPath = Files.createDirectories(workSheetPath.getParent().resolve("media"));
        if (drawingsWriter == null) {
            int id = sheet.getWorkbook().incrementDrawingCounter();
            sheet.getWorkbook().addContentType(new ContentType.Override(Const.ContentType.DRAWINGS, "/xl/drawings/drawing" + id + ".xml"));
            sheet.addRel(new Relationship("../drawings/drawing" + id + ".xml", Const.Relationship.DRAWINGS));
            drawingsWriter = new XMLDrawingsWriter(workSheetPath.getParent().resolve("drawings").resolve("drawing" + id + ".xml"));
        }
        return drawingsWriter;
    }

    /**
     * Download remote resources
     *
     * By default, only HTTP or HTTPS protocols are supported.
     * Use {@code HttpURLConnection} to synchronously download remote resources.
     * For more complex scenarios (asynchronous download, Connection pool, authentication, FTP, etc.), please override this method
     *
     * @param picture {@link Picture} info
     * @param url remote url
     * @throws IOException if I/O error occur.
     */
    public void downloadRemoteResource(Picture picture, String url) throws IOException {
        try {
            // Support http or https
            if (url.charAt(0) == 'h') {
                HttpURLConnection con = (HttpURLConnection) new URL(url).openConnection();
                con.connect();
                InputStream is = con.getInputStream();
                if (is != null) {
                    ByteArrayOutputStream bos = new ByteArrayOutputStream(1 << 18);
                    int i;
                    byte[] bytes = new byte[1 << 18];
                    while ((i = is.read(bytes)) > 0) bos.write(bytes, 0, i);
                    downloadCompleted(picture, bos.toByteArray());
                    return;
                }
                con.disconnect();
            }
            // Ignore others
            downloadCompleted(picture, null);
        } catch (IOException e) {
            LOGGER.error("Download remote resource [{}] error", url, e);
            downloadCompleted(picture, null);
        }
    }

    /**
     * Complete downloading, check the file signatures and set the subscript {@code Picture.idx} to idle
     *
     * @param picture {@link Picture} info
     * @param body thr resource data
     * @throws IOException if I/O error occur.
     */
    public void downloadCompleted(Picture picture, byte[] body) throws IOException {
        LOGGER.debug("completed Row: {} len: {}", picture.row, body != null ? body.length : 0);
        if (body == null || body.length == 0) {
            drawingsWriter.complete(picture);
            return;
        }
        // Test file signatures
        FileSignatures.Signature signature = FileSignatures.test(ByteBuffer.wrap(body));
        if (signature != null && signature.isTrusted()) {
            String name = "image" + picture.id + "." + signature.extension;
            // Store onto disk
            Files.write(mediaPath.resolve(name), body, StandardOpenOption.CREATE_NEW);
            picture.picName = name;
            picture.size = signature.width << 16 | signature.height;
            // Add global contentType
            sheet.getWorkbook().addContentType(new ContentType.Default(signature.contentType, signature.extension));
        } else LOGGER.warn("File types that are not allowed");

        // Setting idle
        drawingsWriter.complete(picture);
    }

    // Write picture
    protected void writePictureDirect(int id, String name, int column, int row, FileSignatures.Signature signature) throws IOException {
        Picture picture = createPicture(column, row);
        picture.id = id;
        picture.picName = name;
        picture.size = signature.width << 16 | signature.height;

        // 实例化drawingsWriter
        if (drawingsWriter == null) createDrawingsWriter();
        // Drawing
        drawingsWriter.drawing(picture);

        // Add global contentType
        sheet.getWorkbook().addContentType(new ContentType.Default(signature.contentType, signature.extension));
    }

    /**
     * Picture constructor
     * You can use this method to add general effects
     *
     * @param column cell column
     * @param row cell row
     * @return {@link Picture}
     */
    protected Picture createPicture(int column, int row) {
        Picture picture = new Picture();
        picture.col = column;
        picture.row = row - 1;
        // 默认位置和大小随单元格一起变动
        picture.toCol = column + 1;
        picture.toRow = row;
//        picture.setPadding(1);
        picture.setPaddingTop(1).setPaddingRight(-1).setPaddingBottom(-1).setPaddingLeft(1);
        picture.effect = getColumn(column).effect;

        return picture;
    }

    /**
     * 收集表头信息，如果有共享字符串标记则初始化SharedStringsTable
     */
    protected void collectHeaderColumns() {
        // 判断是否有共享设置，有共享需要对SharedStrings进行初始化
        boolean hasSharedString = false;
        for (Column col : columns) {
            // 自适应列宽标识
            includeAutoWidth |= col.getAutoSize() == 1;
            hasSharedString |= col.isShare();
        }
        // 初始化SharedStringsTable
        if (hasSharedString && sst != null) sst.init();
        // 如果有自适应列宽则创建临时数组
        if (includeAutoWidth) {
            columnWidths = new double[columns.length];
        }
    }

    /**
     * 按cellXfs下标缓存字体
     */
    protected Font[] fs = new Font[100];
    /**
     * 通过单元格样式索引获取{@code FontMetrics}用以计算文本宽度
     *
     * @param xf 单元格样式索引
     * @return 字体度量对象
     */
    protected Font getFont(int xf) {
        if (xf >= fs.length) fs = Arrays.copyOf(fs, Math.max(xf, fs.length + 100));
        Font f = fs[xf];
        if (f == null) {
            // 通过xf获取当前文本对应的字体
            int style = styles.getStyleByIndex(xf);
            fs[xf] = f = styles.getFont(style);
        }
        return f;
    }

    /**
     * 获取列属性
     *
     * @param index 列下标（从0开始）
     * @return 列属性（不为{@code null}）
     */
    protected Column getColumn(int index) {
        Column hc = index < columns.length ? columns[index] : null;
        if (hc == null) {
            hc = Column.UNALLOCATED_COLUMN;
            hc.colNum = index + sheet.getStartColNum();
        }
        return hc;
    }
}

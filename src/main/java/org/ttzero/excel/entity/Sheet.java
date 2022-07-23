/*
 * Copyright (c) 2017, guanquan.wang@yandex.com All Rights Reserved.
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


import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.ttzero.excel.annotation.TopNS;
import org.ttzero.excel.entity.e7.XMLWorksheetWriter;
import org.ttzero.excel.entity.style.Border;
import org.ttzero.excel.entity.style.BorderStyle;
import org.ttzero.excel.entity.style.Fill;
import org.ttzero.excel.entity.style.Font;
import org.ttzero.excel.entity.style.Horizontals;
import org.ttzero.excel.entity.style.NumFmt;
import org.ttzero.excel.entity.style.Styles;
import org.ttzero.excel.entity.style.Verticals;
import org.ttzero.excel.manager.Const;
import org.ttzero.excel.manager.RelManager;
import org.ttzero.excel.processor.ConversionProcessor;
import org.ttzero.excel.reader.Cell;
import org.ttzero.excel.util.FileUtil;

import java.awt.Color;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.ObjectInputStream;
import java.io.ObjectOutputStream;
import java.lang.reflect.InvocationTargetException;
import java.nio.file.Path;
import java.util.HashMap;
import java.util.Map;

import static org.ttzero.excel.manager.Const.ROW_BLOCK_SIZE;
import static org.ttzero.excel.util.StringUtil.isEmpty;

/**
 * Each worksheet corresponds to one or more sheet.xml of physical.
 * When the amount of data exceeds the upper limit of the worksheet,
 * the extra data will be written in the next worksheet page of the
 * current position, with the name of the parent worksheet. After
 * adding "(1,2,3...n)" as the name of the copied sheet, the pagination
 * is automatic without additional settings.
 * <p>
 * Usually worksheetWriter calls the
 * {@link #nextBlock} method to load a row-block for writing.
 * When the row-block returns the flag EOF, mean is the current worksheet
 * finished written, and the next worksheet is written.
 * <p>
 * Extends the existing worksheet to implement a custom data source worksheet.
 * The data source can be micro-services, Mybatis, JPA or any others. If
 * the data source returns an array of json objects, please convert to
 * an object ArrayList or Map ArrayList, the object ArrayList needs to
 * extends {@link ListSheet}, the Map ArrayList needs to extends
 * {@link ListMapSheet} and implement the {@link ListSheet#more} method.
 * <p>
 * If other formats cannot be converted to ArrayList, you
 * need to inherit from the base class {@link Sheet} and implement the
 * {@link #resetBlockData} and {@link #getHeaderColumns} methods.
 *
 * @see ListSheet
 * @see ListMapSheet
 * @see ResultSetSheet
 * @see StatementSheet
 * @see CSVSheet
 *
 * @author guanquan.wang on 2017/9/26.
 */
@TopNS(prefix = {"", "r"}, value = "worksheet"
        , uri = {Const.SCHEMA_MAIN, Const.Relationship.RELATIONSHIP})
public abstract class Sheet implements Cloneable, Storable {
    protected final Logger LOGGER = LoggerFactory.getLogger(getClass());

    protected Workbook workbook;

    protected String name;
    protected org.ttzero.excel.entity.Column[] columns;
    protected WaterMark waterMark;
    protected RelManager relManager;
    protected int id;
    /**
     * The header column comments
     */
    protected Comments comments;
    /**
     * To mark the cell auto-width
     */
    protected int autoSize;
    /**
     * The default cell width
     */
    protected double width = 20.38D;
    /**
     * The row number
     */
    protected int rows;
    /**
     * Mark the cell is hidden
     */
    protected boolean hidden;

    /**
     * The header style index
     */
    protected int headStyleIndex = -1;

    /**
     * The header style value
     */
    protected int headStyle;

    /**
     * Automatic interlacing color
     */
    protected int autoOdd = -1;
    /**
     * Odd row's background color
     */
    protected int oddFill;
    /**
     * A copy worksheet flag
     */
    protected boolean copySheet;
    protected int copyCount;

    protected RowBlock rowBlock;
    protected IWorksheetWriter sheetWriter;
    /**
     * To mark the header column is ready
     */
    protected boolean headerReady;
    /**
     * Close resource on the last copy worksheet
     */
    protected boolean shouldClose = true;

    protected ICellValueAndStyle cellValueAndStyle;

    /**
     * Force export all attributes
     */
    protected int forceExport;

    /**
     * Ignore header when export
     */
    protected int nonHeader = -1;
    /**
     * Limit row number in worksheet
     */
    private int rowLimit;

    /**
     * Other extend properties
     */
    protected Map<String, Object> extProp = new HashMap<>();
    /**
     * The bit flag of the extended parameter. If there is an extended parameter,
     * the corresponding bit is 1. The lower 16 bits are occupied by the system,
     * and the upper 16 bits can be extended by themselves.
     */
    protected int extPropMark;

    public int getId() {
        return id;
    }

    public void setId(int id) {
        this.id = id;
    }

    public void setSheetWriter(IWorksheetWriter sheetWriter) {
        this.sheetWriter = sheetWriter;
    }

    public void setCellValueAndStyle(ICellValueAndStyle cellValueAndStyle) {
        this.cellValueAndStyle = cellValueAndStyle;
    }

    public Sheet() {
        this(null);
    }

    /**
     * Constructor worksheet
     *
     * @param name the worksheet name
     */
    public Sheet(String name) {
        this.name = name;
        relManager = new RelManager();
    }

    /**
     * Constructor worksheet
     *
     * @param name    the worksheet name
     * @param columns the header info
     */
    public Sheet(String name, final org.ttzero.excel.entity.Column... columns) {
        this(name, null, columns);
    }

    /**
     * Constructor worksheet
     *
     * @param name      the worksheet name
     * @param waterMark the water mark
     * @param columns   the header info
     */
    public Sheet(String name, WaterMark waterMark, final org.ttzero.excel.entity.Column... columns) {
        this.name = name;
        this.columns = columns;
        this.waterMark = waterMark;
        relManager = new RelManager();
    }

    /**
     * Will be deleted soon
     *
     * @deprecated use the new {@link org.ttzero.excel.entity.Column}
     */
    @Deprecated
    public static class Column extends org.ttzero.excel.entity.Column {
        public Column() {
        }

        public Column(String name, Class<?> clazz) {
            super(name, clazz);
        }

        public Column(String name, String key) {
            super(name, key);
        }

        public Column(String name, String key, Class<?> clazz) {
            super(name, key, clazz);
        }

        public Column(String name, Class<?> clazz, ConversionProcessor processor) {
            super(name, clazz, processor);
        }

        public Column(String name, String key, ConversionProcessor processor) {
            super(name, key, processor);
        }

        public Column(String name, Class<?> clazz, boolean share) {
            super(name, clazz, share);
        }

        public Column(String name, String key, boolean share) {
            super(name, key, share);
        }

        public Column(String name, Class<?> clazz, ConversionProcessor processor, boolean share) {
            super(name, clazz, processor, share);
        }

        public Column(String name, String key, Class<?> clazz, ConversionProcessor processor) {
            super(name, key, clazz, processor);
        }

        public Column(String name, String key, ConversionProcessor processor, boolean share) {
            super(name, key, processor, share);
        }

        public Column(String name, Class<?> clazz, int cellStyle) {
            super(name, clazz, cellStyle);
        }

        public Column(String name, String key, int cellStyle) {
            super(name, key, cellStyle);
        }

        public Column(String name, Class<?> clazz, int cellStyle, boolean share) {
            super(name, clazz, cellStyle, share);
        }

        public Column(String name, String key, int cellStyle, boolean share) {
            super(name, key, cellStyle, share);
        }

        /**
         * Setting the cell type
         *
         * @param type the cell type
         * @return the {@link org.ttzero.excel.entity.Column}
         * @deprecated replace it with the {{@link #setNumFmt(String)}} method.
         */
        @Deprecated
        public org.ttzero.excel.entity.Column setType(int type) {
            switch (type) {
                case Const.ColumnType.PARENTAGE:
                    setNumFmt("0.00%_);[Red]-0.00% ");
                    break;
                case Const.ColumnType.RMB:
                    setNumFmt("¥0.00_);[Red]-¥0.00 ");
                    break;
                default:
            }
            return this;
        }
    }

    /**
     * Returns workbook
     *
     * @return current {@link Workbook}
     */
    public Workbook getWorkbook() {
        return workbook;
    }

    /**
     * Setting the workbook
     *
     * @param workbook the {@link Workbook}
     * @return current {@link Sheet}
     */
    public Sheet setWorkbook(Workbook workbook) {
        this.workbook = workbook;
        if (columns != null) {
            for (int i = 0; i < columns.length; i++) {
                columns[i].styles = workbook.getStyles();
            }
        }
        return this;
    }

    /**
     * Output the export detail info
     *
     * @param code the message code in message properties file
     */
    public void what(String code) {
        workbook.what(code);
    }

    /**
     * Output export detail info
     *
     * @param code the message code in message properties file
     * @param args the placeholder values
     */
    public void what(String code, String... args) {
        workbook.what(code, args);
    }

    /**
     * Returns shared string
     *
     * @return global {@link SharedStrings} in workbook
     */
    public SharedStrings getSst() {
        return workbook.getSst();
    }

    /**
     * Return the cell default width
     *
     * @return the width value
     */
    public double getDefaultWidth() {
        return width;
    }

    /**
     * Setting auto resize cell's width
     *
     * @return current {@link Sheet}
     */
    public Sheet autoSize() {
        this.autoSize = 1;
        return this;
    }

    /**
     * Setting fix column width
     *
     * @return current {@link Sheet}
     */
    public Sheet fixSize() {
        this.autoSize = 2;
        return this;
    }

    /**
     * Setting fix column width
     *
     * @param width the column width
     * @return current {@link Sheet}
     */
    public Sheet fixSize(double width) {
        this.autoSize = 2;
        this.width = width;
        if (headerReady) {
            for (org.ttzero.excel.entity.Column hc : columns) {
                hc.setWidth(width);
            }
        }
        return this;
    }

    /**
     * Returns the re-size setting
     *
     * @return 1: auto-size 2:fix-size
     */
    public int getAutoSize() {
        return autoSize;
    }

    /**
     * Test is auto size column width
     *
     * @return true if auto-size
     */
    public boolean isAutoSize() {
        return autoSize == 1;
    }

    /**
     * Cancel the odd row's fill style
     *
     * @return current {@link Sheet}
     */
    public Sheet cancelOddStyle() {
        this.autoOdd = 1;
        return this;
    }

    /**
     * Returns auto setting odd background flag
     *
     * @return 1: auto setting, others none
     */
    public int getAutoOdd() {
        return autoOdd;
    }

    /**
     * Setting auto setting odd background flag
     *
     * @param autoOdd 1: setting, others none
     * @return current {@link Sheet}
     */
    public Sheet setAutoOdd(int autoOdd) {
        this.autoOdd = autoOdd;
        return this;
    }

    /**
     * Setting the odd row's fill style
     *
     * @param fill the fill style
     * @return current {@link Sheet}
     */
    public Sheet setOddFill(Fill fill) {
        this.oddFill = workbook.getStyles().addFill(fill);
        return this;
    }

    /**
     * Returns the odd columns fill style
     *
     * @return the fill style value
     */
    public int getOddFill() {
        return oddFill;
    }

    /**
     * Returns the worksheet name
     *
     * @return the worksheet name
     */
    public String getName() {
        return name;
    }

    /**
     * Setting the worksheet name
     *
     * @param name the worksheet name
     * @return current {@link Sheet}
     */
    public Sheet setName(String name) {
        this.name = name;
        return this;
    }

    /**
     * Returns the header column {@link Comments}
     *
     * @return Columns instance if exists
     */
    public Comments getComments() {
        return comments;
    }

    /**
     * Create a {@link Comments} and add relationship
     *
     * @return a comment instance
     */
    public Comments createComments() {
        if (comments == null) {
            comments = new Comments(id, workbook.getCreator());
            // FIXME Removed at excel version 2013
            addRel(new Relationship("../drawings/vmlDrawing" + id + Const.Suffix.VML, Const.Relationship.VMLDRAWING));

            addRel(new Relationship("../comments" + id + Const.Suffix.XML, Const.Relationship.COMMENTS));
        }
        return comments;
    }

    /**
     * Returns the header column info
     * <p>
     * The copy sheet will use the parent worksheet header information.
     * <p>
     * Use the method {@link #getAndSortHeaderColumns()} to get Columns
     *
     * @return array of column
     */
    protected org.ttzero.excel.entity.Column[] getHeaderColumns() {
        if (!headerReady) {
            if (columns == null) {
                columns = new org.ttzero.excel.entity.Column[0];
            }
        }
        return columns;
    }

    /**
     * Sort column by {@code colIndex}
     *
     * @return header columns
     */
    public org.ttzero.excel.entity.Column[] getAndSortHeaderColumns() {
        if (!headerReady) {
            this.columns = getHeaderColumns();
            // Reset Common Properties
            resetCommonProperties(columns);
            // Sort column index
            sortColumns(columns);
            // Turn to one-base
            for (int i = 0; i < columns.length; i++) {
                if (i > 0 && columns[i - 1].colIndex >= columns[i].colIndex) columns[i].colIndex = columns[i - 1].colIndex + 1;
                else if (columns[i].colIndex <= i) columns[i].colIndex = i + 1;
                else columns[i].colIndex++;
            }

            // Check the limit of columns
            checkColumnLimit();
            headerReady |= (this.columns.length > 0);

            // Mark ext-properties
            markExtProp();
        }
        return columns;
    }

    protected void resetCommonProperties(org.ttzero.excel.entity.Column[] columns) {
        for (org.ttzero.excel.entity.Column column : columns) {
            if (column == null) continue;
            if (column.styles == null) column.styles = workbook.getStyles();
        }
    }

    protected void sortColumns(org.ttzero.excel.entity.Column[] columns) {
        if (columns.length <= 1) return;
        int j = 0;
        for (int i = 0; i < columns.length; i++) {
            if (columns[i].colIndex >= 0) {
                int n = search(columns, j, columns[i].colIndex);
                if (n < i) insert(columns, n, i);
                j++;
            }
        }
        // Finished
        if (j == columns.length) return;
        int n = columns[0].colIndex;
        for (int i = 0; i < columns.length && j < columns.length; ) {
            if (n > i) {
                for (int k = Math.min(n - i, columns.length - j); k > 0; k--, j++)
                    insert(columns, i++, j);
            } else i++;
            if (i < columns.length) n = columns[i].colIndex;
        }
    }

    protected int search(org.ttzero.excel.entity.Column[] columns, int n, int k) {
        int i = 0;
        for (; i < n && columns[i].colIndex <= k; i++) ;
        return i;
    }

    private void insert(org.ttzero.excel.entity.Column[] columns, int n, int k) {
        org.ttzero.excel.entity.Column t = columns[k];
        System.arraycopy(columns, n, columns, n + 1, k - n);
        columns[n] = t;
    }

    /**
     * Setting the header rows's columns
     *
     * @param columns the header row's columns
     * @return current {@link Sheet}
     */
    public Sheet setColumns(final org.ttzero.excel.entity.Column[] columns) {
        this.columns = columns;
        return this;
    }

    /**
     * Returns the {@link WaterMark}
     *
     * @return the {@link WaterMark} in worksheet
     * @see WaterMark
     */
    public WaterMark getWaterMark() {
        return waterMark;
    }

    /**
     * Setting the {@link WaterMark}
     *
     * @param waterMark the {@link WaterMark}
     * @return current {@link Sheet}
     */
    public Sheet setWaterMark(WaterMark waterMark) {
        this.waterMark = waterMark;
        return this;
    }

    /**
     * Returns the worksheet is hidden
     *
     * @return true: hidden, false: not hidden
     */
    public boolean isHidden() {
        return hidden;
    }

    /**
     * Setting the worksheet status
     *
     * @return current {@link Sheet}
     */
    public Sheet hidden() {
        this.hidden = true;
        return this;
    }

    /**
     * Force export of attributes without {@link org.ttzero.excel.annotation.ExcelColumn} annotations
     *
     * @return current {@link Sheet}
     */
    public Sheet forceExport() {
        this.forceExport = 1;
        return this;
    }

    /**
     * Cancel force export
     *
     * @return current {@link Sheet}
     */
    public Sheet cancelForceExport() {
        this.forceExport = 2;
        return this;
    }

    /**
     * Returns the force export
     *
     * @return 1 if force, otherwise returns 0
     */
    public int getForceExport() {
        return forceExport;
    }

    /**
     * abstract method close
     *
     * @throws IOException if I/O error occur
     */
    public void close() throws IOException {
        if (sheetWriter != null) {
            sheetWriter.close();
        }
    }

    /**
     * Write worksheet data to path
     *
     * @param path the storage path
     * @throws IOException if I/O error occur
     */
    @Override
    public void writeTo(Path path) throws IOException {
        if (sheetWriter == null) {
            throw new ExcelWriteException("Worksheet writer is not instanced.");
        }
        if (!headerReady) {
            getAndSortHeaderColumns();
        }
        if (rowBlock == null) {
            rowBlock = new RowBlock(getRowBlockSize());
        } else rowBlock.reopen();

        if (!copySheet) {
            paging();
        }

        sheetWriter.writeTo(path);
    }

    /**
     * Split worksheet data
     */
    protected void paging() { }

    /**
     * Add relationship
     *
     * @param rel Relationship
     * @return current worksheet
     */
    public Sheet addRel(Relationship rel) {
        relManager.add(rel);
        return this;
    }

    public Relationship findRel(String key) {
        return relManager.likeByTarget(key);
    }

    /**
     * Returns the worksheet name
     *
     * @return name of worksheet
     */
    public String getFileName() {
        return "sheet" + id + cellValueAndStyle.getFileSuffix();
    }

    /**
     * Setting the header column styles
     *
     * @param font   the font
     * @param fill   the fill style
     * @param border the border style
     * @return current {@link Sheet}
     */
    public Sheet setHeadStyle(Font font, Fill fill, Border border) {
        return setHeadStyle(null, font, fill, border, Verticals.CENTER, Horizontals.CENTER);
    }

    /**
     * Setting the header column styles
     *
     * @param font       the font
     * @param fill       the fill style
     * @param border     the border style
     * @param vertical   the vertical style
     * @param horizontal the horizontal style
     * @return current {@link Sheet}
     */
    public Sheet setHeadStyle(Font font, Fill fill, Border border, int vertical, int horizontal) {
        return setHeadStyle(null, font, fill, border, vertical, horizontal);
    }

    /**
     * Setting the header column styles
     *
     * @param numFmt     the number format
     * @param font       the font
     * @param fill       the fill style
     * @param border     the border style
     * @param vertical   the vertical style
     * @param horizontal the horizontal style
     * @return current {@link Sheet}
     */
    public Sheet setHeadStyle(NumFmt numFmt, Font font, Fill fill, Border border, int vertical, int horizontal) {
        Styles styles = workbook.getStyles();
        headStyle = (numFmt != null ? styles.addNumFmt(numFmt) : 0)
            | (font != null ? styles.addFont(font) : 0)
            | (fill != null ? styles.addFill(fill) : 0)
            | (border != null ? styles.addBorder(border) : 0)
            | vertical
            | horizontal;
        headStyleIndex = styles.of(headStyle);
        return this;
    }

    /**
     * Setting the header cell styles
     *
     * @param style the styles value
     * @return current {@link Sheet}
     */
    public Sheet setHeadStyle(int style) {
        headStyle = style;
        headStyleIndex = workbook.getStyles().of(style);
        return this;
    }

    /**
     * Setting the header cell styles
     *
     * @param styleIndex the styles index
     * @return current {@link Sheet}
     */
    public Sheet setHeadStyleIndex(int styleIndex) {
        headStyleIndex = styleIndex;
        headStyle = workbook.getStyles().getStyleByIndex(styleIndex);
        return this;
    }

    /**
     * Returns the header style value
     *
     * @return 0 if not set
     */
    public int getHeadStyle() {
        return headStyle;
    }

    /**
     * Returns the header style index
     *
     * @return -1 if not set
     */
    public int getHeadStyleIndex() {
        return headStyleIndex;
    }

    /**
     * Custom header style according to parameters
     *
     * @param fontColor the font color
     * @param fillBgColor the fill background color
     * @return style value
     */
    public int buildHeadStyle(String fontColor, String fillBgColor) {
        Styles styles = workbook.getStyles();
        Font font = new Font(workbook.getI18N().getOrElse("local-font-family", "Arial")
                , 12, Font.Style.BOLD, Styles.toColor(fontColor));
        return styles.addFont(font)
                | styles.addFill(Fill.parse(fillBgColor))
                | styles.addBorder(new Border().setBorder(BorderStyle.THIN, new Color(191, 191, 191)))
                | Verticals.CENTER
                | Horizontals.CENTER;
    }

    /**
     * Build default header style
     *
     * @return style value
     */
    public int defaultHeadStyle() {
        return headStyle != 0 ? headStyle : (headStyle = this.buildHeadStyle("#ffffff", "#666699"));
    }

    /**
     * Build default header style
     *
     * @return style index
     */
    public int defaultHeadStyleIndex() {
        if (headStyleIndex == -1) {
            setHeadStyle(this.buildHeadStyle("#ffffff", "#666699"));
        }
        return headStyleIndex;
    }

    protected static boolean nonOrIntDefault(int style) {
        return style == -1
            || style == Styles.defaultIntBorderStyle()
            || style == Styles.defaultIntStyle();
    }

    /**
     * Returns total rows in this worksheet
     *
     * @return -1 if unknown or uncertain
     */
    public int size() {
        return -1;
    }

    /**
     * Returns a row-block. The row-block is content by 32 rows
     *
     * @return a row-block
     */
    public RowBlock nextBlock() {
        // clear first
        rowBlock.clear();

        if (columns.length > 0) {
            resetBlockData();
        }

        return rowBlock.flip();
    }

    /**
     * The worksheet is written by units of row-block. The default size
     * of a row-block is 32, which means that 32 rows of data are
     * written at a time. If the data is not enough, the {@code more()}
     * method will be called to get more data.
     *
     * @return the row-block size
     */
    public int getRowBlockSize() {
        return ROW_BLOCK_SIZE;
    }

    /**
     * Write some final info
     *
     * @param workSheetPath the worksheet path
     * @throws IOException if I/O error occur
     */
    public void afterSheetAccess(Path workSheetPath) throws IOException {
        // relationship
        if (sheetWriter instanceof XMLWorksheetWriter) {
            relManager.write(workSheetPath, getFileName());
        }

        // others ...
    }

    /**
     * Returns the copy worksheet name
     *
     * @return the name of copy worksheet
     */
    protected String getCopySheetName() {
        int sub = copyCount;
        String _name = name;
        // reset name
        int i = name.lastIndexOf('(');
        if (i > 0) {
            int fs = Integer.parseInt(name.substring(i + 1, name.lastIndexOf(')')));
            _name = name.substring(0, name.charAt(i - 1) == ' ' ? i - 1 : i);
            if (++fs > sub) sub = fs;
        }
        return _name + " (" + (sub) + ")";
    }

    @Override
    public Sheet clone() {
        Sheet copy = null;
        try {
            copy = (Sheet) super.clone();
        } catch (CloneNotSupportedException e) {
            ObjectOutputStream oos = null;
            ObjectInputStream ois = null;
            try {
                ByteArrayOutputStream bos = new ByteArrayOutputStream();
                oos = new ObjectOutputStream(bos);
                oos.writeObject(this);

                ois = new ObjectInputStream(new ByteArrayInputStream(bos.toByteArray()));
                copy = (Sheet) ois.readObject();
            } catch (IOException | ClassNotFoundException e1) {
                try {
                    copy = getClass().getConstructor().newInstance();
                } catch (NoSuchMethodException | IllegalAccessException
                    | InstantiationException | InvocationTargetException e2) {
                    e2.printStackTrace();
                }
            } finally {
                FileUtil.close(oos);
                FileUtil.close(ois);
            }
        }
        if (copy != null) {
            copy.copyCount = ++copyCount;
            copy.name = getCopySheetName();
            copy.relManager = relManager.deepClone();
            copy.sheetWriter = sheetWriter.clone().setWorksheet(copy);
            copy.copySheet = true;
            copy.rows = 0;
        }
        return copy;
    }

    /**
     * Check the limit of columns
     */
    public void checkColumnLimit() {
        int a = columns.length > 0 ? columns[columns.length - 1].colIndex : 0
            , b = sheetWriter.getColumnLimit();
        if (a > b) {
            throw new TooManyColumnsException(a, b);
        } else {
            boolean noneHeader = columns == null || columns.length == 0;
            if (!noneHeader) {
                int n = 0;
                for (org.ttzero.excel.entity.Column column : columns) {
                    if (isEmpty(column.name)) n++;
                }
                noneHeader = n == columns.length;
            }
            if (noneHeader) {
                if (rows > 0) rows--;
                ignoreHeader();
            } else this.nonHeader = 0;
            this.rowLimit = sheetWriter.getRowLimit() - (this.nonHeader ^ 1);
        }
    }

    /**
     * Check the header information is exist
     *
     * @return true if exist
     */
    public boolean hasHeaderColumns() {
        return columns != null && columns.length > 0;
    }

    /**
     * Int conversion to column string number.
     * The max column on sheet is {@code 16_384} after office 2007 and {@code 256} in office 2003
     * <blockquote><pre>
     * int    | column number
     * -------|---------
     * 1      | A
     * 10     | J
     * 26     | Z
     * 27     | AA
     * 28     | AB
     * 53     | BA
     * 16_384 | XFD
     * </pre></blockquote>
     * @param n the column number
     * @return column string
     */
    public static char[] int2Col(int n) {
        char[][] cache_col = cache.get();
        char[] c;
        char A = 'A';
        if (n <= 26) {
            c = cache_col[0];
            c[0] = (char) (n - 1 + A);
        } else if (n <= 702) {
            int t = n / 26, w = n % 26;
            if (w == 0) {
                t--;
                w = 26;
            }
            c = cache_col[1];
            c[0] = (char) (t - 1 + A);
            c[1] = (char) (w - 1 + A);
        } else {
            int tt = n / 26, t = tt / 26, w = n % 26, m = tt % 26;
            if (w == 0) {
                m--;
                w = 26;
            }
            if (m <= 0) {
                t--;
                m += 26;
            }
            c = cache_col[2];
            c[0] = (char) (t - 1 + A);
            c[1] = (char) (m - 1 + A);
            c[2] = (char) (w - 1 + A);
        }
        return c;
    }

    private static final ThreadLocal<char[][]> cache
        = ThreadLocal.withInitial(() -> new char[][]{ {65}, {65, 65}, {65, 65, 65} });

//    /**
//     * Check empty header row
//     *
//     * @return true if none header row
//     */
//    public boolean hasNonHeader() {
//        int nonHeader = getNonHeader();
//        if (nonHeader == -1) {
//            columns = getAndSortHeaderColumns();
//            boolean noneHeader = columns == null || columns.length == 0;
//            if (!noneHeader) {
//                int n = 0;
//                for (org.ttzero.excel.entity.Column column : columns) {
//                    if (isEmpty(column.name)) n++;
//                }
//                noneHeader = n == columns.length;
//            }
//            if (noneHeader) {
////                rows--;
//                ignoreHeader();
//            } else this.nonHeader = 0;
//            return noneHeader;
//        }
//        return nonHeader == 1;
//    }

    /**
     * Settings nonHeader property
     *
     * @return current Worksheet
     */
    public Sheet ignoreHeader() {
        this.nonHeader = 1;
        return this;
    }

    /**
     * Returns the nonHeader value.
     *
     * @return -1, 0, 1 means not-set, include header, exclude header
     */
    public int getNonHeader() {
        return nonHeader;
    }

    /**
     * The Worksheet row limit
     *
     * @return the limit
     */
    protected int getRowLimit() {
        return rowLimit;
    }

    /**
     * Append extend property
     *
     * @param key key with which the specified value is to be associated
     * @param value value to be associated with the specified key
     * @return current Worksheet
     */
    public Sheet putExtProp(String key, Object value) {
        extProp.put(key, value);
        return this;
    }

    /**
     * If the specified key is not already associated with a value (or is mapped
     * to {@code null}) associates it with the given value and returns
     * {@code null}, else returns the current value.
     *
     * @param key key with which the specified value is to be associated
     * @param value value to be associated with the specified key
     * @return current Worksheet
     */
    public Sheet putExtPropIfAbsent(String key, Object value) {
        extProp.putIfAbsent(key, value);
        return this;
    }

    /**
     * Copies all of the mappings from the specified map to extend properties
     *
     * @param m mappings to be stored in this map
     * @return current Worksheet
     */
    public Sheet putAllExtProp(Map<String, Object> m) {
        extProp.putAll(m);
        return this;
    }

    /**
     * Returns the value to which the specified key in extend property,
     * or {@code null} if it contains no mapping for the key.
     *
     * @param key the key whose associated value is to be returned
     * @return the extend property value
     */
    public Object getExtPropValue(String key) {
        return extProp.get(key);
    }

    /**
     * Shallow copy all extend properties
     *
     * @return all extend properties
     */
    public Map<String, Object> getExtPropAsMap() {
        return new HashMap<>(extProp);
    }

    /**
     * Mark ext-properties bit
     */
    protected void markExtProp() {
        // Mark Freeze Panes
        extPropMark |= getExtPropValue(Const.ExtendPropertyKey.FREEZE) != null ? 1 : 0;
        // Mark global style design
        extPropMark |= getExtPropValue(Const.ExtendPropertyKey.STYLE_DESIGN) != null ? 1 << 1 : 0;
    }

    ////////////////////////////Abstract function\\\\\\\\\\\\\\\\\\\\\\\\\\\

    /**
     * Each row-block is multiplexed and will be called to reset
     * the data when a row-block is completely written.
     * Call the {@link #getRowBlockSize()} method to get
     * the row-block size, call the {@link ICellValueAndStyle#reset(int, Cell, Object, org.ttzero.excel.entity.Column)}
     * method to set value and styles.
     */
    protected abstract void resetBlockData();
}

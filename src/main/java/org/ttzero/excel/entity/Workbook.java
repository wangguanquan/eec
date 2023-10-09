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

import org.ttzero.excel.entity.csv.CSVWorkbookWriter;
import org.ttzero.excel.entity.e7.ContentType;
import org.ttzero.excel.entity.e7.XMLWorkbookWriter;
import org.ttzero.excel.entity.style.Fill;
import org.ttzero.excel.entity.style.PatternType;
import org.ttzero.excel.entity.style.Styles;
import org.ttzero.excel.manager.docProps.Core;
import org.ttzero.excel.util.FileUtil;
import org.ttzero.excel.util.StringUtil;

import java.awt.Color;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Arrays;
import java.util.function.BiConsumer;

import static org.ttzero.excel.util.FileUtil.exists;

/**
 * The workbook is the basic unit of Excel, and an 'xlsx' or 'xls' file
 * corresponds to a workbook instance.
 * <p>
 * When writing an Excel file, you must setting the property first and
 * then add {@link Sheet} into Workbook, finally call the {@link #writeTo}
 * method to perform the write operation. The default file format is open-xml(xlsx).
 * <p>
 * The property contains {@link #setName(String)}, {@link #setCreator(String)},
 * {@link #setCompany(String)}, {@link #setAutoSize(boolean)} and {@link #setZebraLine(Fill)}
 * You can also call {@link #setWorkbookWriter(IWorkbookWriter)} method to setting
 * a custom WorkbookWriter to achieve special demand.
 * <p>
 * The {@link #writeTo} method is a terminating statement, and all settings
 * placed after this statement will not be reflected in the final Excel file.
 * <p>
 * A typical example as follow:
 * <blockquote><pre>
 * new Workbook("{name}", "{author}")
 *     // Auto size the column width
 *     .setAutoSize(true)
 *     // Add a Worksheet
 *     .addSheet(new ListSheet&lt;Item&gt;("{worksheet name}").setData(new ArrayList&lt;&gt;()))
 *     // Add an other Worksheet
 *     .addSheet(new ListMapSheet("{worksheet name}").setData(new ArrayList&lt;&gt;()))
 *     // Write to absolute path '/tmp/{name}.xlsx'
 *     .writeTo(Paths.get("/tmp/"));</pre></blockquote>
 * <p>Some referer links:
 * <a href="https://poi.apache.org">POI</a>&nbsp;|&nbsp;
 * <a href="https://msdn.microsoft.com/library">Office 365</a>&nbsp;|&nbsp;
 * <a href="https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet(v=office.14).aspx#">DocumentFormat.OpenXml.Spreadsheet Namespace</a>&nbsp;|&nbsp;
 * <a href="https://docs.microsoft.com/zh-cn/previous-versions/office/office-12/ms406049(v=office.12)">介绍 Microsoft Office (2007) Open XML 文件格式</a>
 *
 * @author guanquan.wang on 2017/9/26.
 */
public class Workbook implements Storable {
    /**
     * The Workbook name, reaction to the Excel file name
     */
    private String name;
    private Sheet[] sheets;
    private WaterMark waterMark;
    private int size;
    @Deprecated
    private Connection con;
    /**
     * Auto size flag
     */
    private boolean autoSize;
    /**
     * Author
     */
    private String creator;
    private Core core;
    /**
     * Specify a company name(null able)
     */
    private String company;
    /**
     * The zebra-line fill style
     */
    private Fill zebraFill;
    /**
     * A progress window
     */
    private BiConsumer<Sheet, Integer> progressConsumer;
    private final I18N i18N;

    private SharedStrings sst;
    private Styles styles;

    private IWorkbookWriter workbookWriter;

    /**
     * Force export all attributes
     */
    private int forceExport;
    /**
     * A global ContentType attributes
     */
    private final ContentType contentType;
    /**
     * Drawing worksheet counter
     */
    private int drawingCounter;
    /**
     * Count of media in global workbook
     */
    private int mediaCounter;

    /**
     * Create a unnamed workbook
     *
     * EEC finds the 'non-name-file' keyword under the {@code resources/I18N/message.XXX.properties}
     * file first. If there contains the keyword, use this value as the default file name,
     * otherwise use 'Non name' as the name.
     */
    public Workbook() {
        this(null);
    }

    /**
     * Create a workbook with the specified name. Use this name as
     * the file name when saving to disk
     *
     * @param name the workbook name
     */
    public Workbook(String name) {
        this(name, null);
    }

    /**
     * Create a workbook with the specified name and author.
     *
     * @param name    the workbook name
     * @param creator the author, it will getting the
     *      {@code System.getProperty("user.name")} if it not be setting.
     */
    public Workbook(String name, String creator) {
        this.name = name;
        this.creator = creator;
        sheets = new Sheet[3]; // Create three worksheet
        i18N = new I18N();
        contentType = new ContentType();
    }

    /**
     * Returns the workbook name
     *
     * @return the workbook name
     */
    public String getName() {
        return name;
    }

    /**
     * Setting the workbook name
     *
     * @param name the workbook name
     * @return the {@link Workbook}
     */
    public Workbook setName(String name) {
        this.name = name;
        return this;
    }

    /**
     * Returns the autoOdd setting
     *
     * @return 1 if odd-fill
     * @deprecated replace with {@code getZebraFill() != null}
     */
    @Deprecated
    public int getAutoOdd() {
        return hasZebraFill() ? 1 : 0;
    }

    /**
     * Returns the excel author
     *
     * @return the author
     */
    public String getCreator() {
        return creator;
    }

    /**
     * Returns the company name where the author is employed
     *
     * @return the company name
     */
    public String getCompany() {
        return company;
    }

    /**
     * Returns the odd-fill style
     *
     * @return the {@link Fill} style
     * @deprecated rename to {@link #getZebraFill()}
     */
    @Deprecated
    public Fill getOddFill() {
        return getZebraFill();
    }

    /**
     * Returns the {@link I18N} util
     *
     * @return the {@link I18N} util
     */
    public I18N getI18N() {
        return i18N;
    }

    /**
     * Returns the size of {@link Sheet} in this workbook
     *
     * @return ths size of Worksheet
     */
    public int getSize() {
        return size;
    }

    /**
     * Returns the basic information about workbook
     *
     * @return the {@link Core} instance
     */
    public Core getCore() {
        return core;
    }

    /**
     * Setting basic information,such as title, subject, keyword, category...
     *
     * @param core the {@link Core} instance
     * @return ths size of Worksheet
     */
    public Workbook setCore(Core core) {
        this.core = core;
        return this;
    }

    /**
     * Returns the Shared String Table
     *
     * @return the global {@link SharedStrings}
     */
    public SharedStrings getSst() {
        // CSV do not need SharedStringTable
        if (!(workbookWriter instanceof CSVWorkbookWriter) && sst == null)
            sst = new SharedStrings();
        return sst;
    }

    /**
     * Returns all {@link Sheet} in this workbook
     *
     * @return array of {@link Sheet}
     */
    public final Sheet[] getSheets() {
        return Arrays.copyOf(sheets, size);
    }

    /**
     * Returns the {@link WaterMark}
     *
     * @return the {@link Workbook}
     */
    public WaterMark getWaterMark() {
        return waterMark;
    }

    /**
     * Setting {@link WaterMark}
     * <p>Use {@link WaterMark#of} method to create a water mark.</p>
     *
     * @param waterMark the water mark
     * @return the {@link Workbook}
     */
    public Workbook setWaterMark(WaterMark waterMark) {
        this.waterMark = waterMark;
        return this;
    }

    /**
     * Setting the database {@link Connection}
     * <p>
     * EEC does not actively close the database connection,
     * and needs to be manually closed externally. The {@link java.sql.Statement}
     * and {@link ResultSet} generated inside this EEC will
     * be actively closed.
     *
     * @param con the database connection
     * @return the {@link Workbook}
     * @deprecated insecurity
     */
    @Deprecated
    public Workbook setConnection(Connection con) {
        this.con = con;
        return this;
    }

    /**
     * Setting auto-adjust the column width
     *
     * @param autoSize boolean value
     * @return the {@link Workbook}
     */
    public Workbook setAutoSize(boolean autoSize) {
        this.autoSize = autoSize;
        return this;
    }

    /**
     * Returns whether to auto-adjust the column width
     *
     * @return true if auto-adjust the column width
     */
    public boolean isAutoSize() {
        return autoSize;
    }

    /**
     * Force export of attributes without {@link org.ttzero.excel.annotation.ExcelColumn} annotations
     *
     * @return the {@link Workbook}
     */
    public Workbook forceExport() {
        this.forceExport = 1;
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
     * Returns the global {@link Styles}
     *
     * @return the Styles
     */
    public Styles getStyles() {
        // CSV do not need Styles
        if (!(workbookWriter instanceof CSVWorkbookWriter) && styles == null)
            styles = Styles.create(i18N);
        return styles;
    }

    /**
     * Specify a custom global {@link Styles}
     *
     * @param styles custom Styles
     * @return the {@link Workbook}
     */
    public Workbook setStyles(Styles styles) {
        this.styles = styles;
        return this;
    }

    /**
     * Setting the excel author name.
     * <p>
     * If you do not set the creator it will get the current OS login user name,
     * usually this is not a good idea. Applications are usually publish on server
     * or cloud server, getting the system login user name doesn't make sense.
     * If you don't want to set it and default setting the system login user name,
     * you can set it to an empty string ("").
     * <p>
     * Does anyone agree with this idea? If anyone agrees, I will consider removing
     * this setting.
     *
     * @param creator the author name
     * @return the {@link Workbook}
     */
    public Workbook setCreator(String creator) {
        this.creator = creator;
        return this;
    }

    /**
     * Setting the name of the company where the author is employed
     *
     * @param company the company name
     * @return the {@link Workbook}
     */
    public Workbook setCompany(String company) {
        this.company = company;
        return this;
    }

    /**
     * Cancel the odd-fill style
     *
     * @return the {@link Workbook}
     * @deprecated rename to {@link #cancelZebraLine()}
     */
    @Deprecated
    public Workbook cancelOddFill() {
        return cancelZebraLine();
    }

    /**
     * Setting the odd-fill style, default fill color is #E2EDDA
     *
     * @param fill the {@link Fill} style
     * @return the {@link Workbook}
     * @deprecated rename to {@link #setZebraLine(Fill)}
     */
    @Deprecated
    public Workbook setOddFill(Fill fill) {
        return setZebraLine(fill);
    }

    /**
     * Setting the zebra-line fill style
     *
     * @param fill the zebra-line {@link Fill} style
     * @return the {@link Workbook}
     */
    public Workbook setZebraLine(Fill fill) {
        this.zebraFill = fill;
        return this;
    }

    /**
     * Cancel the zebra-line style
     *
     * @return the {@link Workbook}
     */
    public Workbook cancelZebraLine() {
        this.zebraFill = null;
        return this;
    }

    /**
     * Setting zebra-line style, the default fill color is #EFF5EB
     *
     * @return the {@link Workbook}
     */
    public Workbook defaultZebraLine() {
        return setZebraLine(new Fill(PatternType.solid, new Color(233, 234, 236)));
    }

    /**
     * Returns the zebra-line fill style
     *
     * @return the {@link Fill} style
     */
    public Fill getZebraFill() {
        return zebraFill;
    }

    /**
     * Returns current workbook has zebra-line
     *
     * @return true if zebra-line is not null
     */
    public boolean hasZebraFill() {
        return zebraFill != null && zebraFill.getPatternType() != PatternType.none;
    }

    /**
     * Returns the {@link IWorkbookWriter}
     *
     * @return the workbook writer
     */
    public IWorkbookWriter getWorkbookWriter() {
        if (workbookWriter == null)
            workbookWriter = new XMLWorkbookWriter(this);
        return workbookWriter;
    }

    /**
     * Add a {@link Sheet} to the tail
     *
     * @param sheet a Worksheet
     * @return the {@link Workbook}
     */
    public Workbook addSheet(Sheet sheet) {
        ensureCapacityInternal();
        sheet.setWorkbook(this);
        sheets[size++] = sheet;
        return this;
    }

    /**
     * Add a {@link ListSheet} to the tail with header {@link Column} setting.
     * Also you can use {@code addSheet(new ListSheet&lt;&gt;(data, columns)}
     * to achieve the same effect.
     *
     * @param data    List&lt;?&gt; data
     * @param columns the header columns
     * @return the {@link Workbook}
     * @deprecated use {@link #addSheet(Sheet)}
     */
    @Deprecated
    public Workbook addSheet(List<?> data, Column... columns) {
        return addSheet(null, data, columns);
    }

    /**
     * Add a {@link ListSheet} to the tail with Worksheet name
     * and header {@link Column} setting. Also you can use
     * {@code addSheet(new ListSheet&lt;&gt;(name, data, columns)}
     * to achieve the same effect.
     *
     * @param name    the name of worksheet
     * @param data    List&lt;?&gt; data
     * @param columns the header columns
     * @return the {@link Workbook}
     * @deprecated use {@link #addSheet(Sheet)}
     */
    @Deprecated
    @SuppressWarnings({"unchecked", "rawtypes"})
    public Workbook addSheet(String name, List<?> data, Column... columns) {
        Object o;
        if (data == null || data.isEmpty() || (o = getFirst(data)) == null) {
            addSheet(new EmptySheet(name, columns));
            return this;
        }

        if (o instanceof Map) {
            addSheet(new ListMapSheet(name, columns).setData((List<Map<String, ?>>) data));
        } else {
            addSheet(new ListSheet(name, columns).setData(data));
        }
        return this;
    }

    // Find the first not null data
    private Object getFirst(List<?> data) {
        if (data == null || data.isEmpty()) return null;
        Object first = data.get(0);
        if (first != null) return first;
        int i = 1;
        do {
            first = data.get(i++);
        } while (first == null);
        return first;
    }

    /**
     * Add a {@link ResultSetSheet} to the tail with header {@link Column} setting.
     * Also you can use {@code addSheet(new ResultSetSheet(rs, columns)}
     * to achieve the same effect.
     *
     * @param rs      the {@link ResultSet}
     * @param columns the header columns
     * @return the {@link Workbook}
     * @deprecated use {@link #addSheet(Sheet)}
     */
    @Deprecated
    public Workbook addSheet(ResultSet rs, Column... columns) {
        return addSheet(null, rs, columns);
    }

    /**
     * Add a {@link ResultSetSheet} to the tail with worksheet name
     * and header {@link Column} setting. Also you can use
     * {@code addSheet(new ResultSetSheet(name, rs, columns)}
     *
     * @param name    the worksheet name
     * @param rs      the {@link ResultSet}
     * @param columns the header columns
     * @return the {@link Workbook}
     * @deprecated use {@link #addSheet(Sheet)}
     */
    @Deprecated
    public Workbook addSheet(String name, ResultSet rs, Column... columns) {
        ResultSetSheet sheet = new ResultSetSheet(name, columns);
        sheet.setRs(rs);
        addSheet(sheet);
        return this;
    }

    /**
     * Add a {@link StatementSheet} to the tail with header {@link Column} setting.
     * Also you can use {@code addSheet(new StatementSheet(connection, sql, columns)}
     * to achieve the same effect.
     *
     * @param sql     the query SQL string
     * @param columns the header columns
     * @return the {@link Workbook}
     * @throws SQLException if a database access error occurs
     * @deprecated use {@link #addSheet(Sheet)}
     */
    @Deprecated
    public Workbook addSheet(String sql, Column... columns) throws SQLException {
        return addSheet(null, sql, columns);
    }

    /**
     * Add a {@link StatementSheet} to the tail with worksheet name
     * and header {@link Column} setting. Also you can use
     * {@code addSheet(new StatementSheet(name, connection, sql, columns)}
     * to achieve the same effect.
     *
     * @param name    the worksheet name
     * @param sql     the query SQL string
     * @param columns the header columns
     * @return the {@link Workbook}
     * @throws SQLException if a database access error occurs
     * @deprecated use {@link #addSheet(Sheet)}
     */
    @Deprecated
    public Workbook addSheet(String name, String sql, Column... columns) throws SQLException {
        PreparedStatement ps = con.prepareStatement(sql
            , ResultSet.TYPE_FORWARD_ONLY, ResultSet.CONCUR_READ_ONLY);
        return addSheet(name, ps, null, columns);
    }

    /**
     * Add a {@link StatementSheet} to the tail with header {@link Column}
     * setting. The {@link ParamProcessor} is a sql parameter replacement
     * function-interface to replace "?" in the sql string.
     * <p>
     * Also you can use {@code addSheet(new StatementSheet(connection, sql, paramProcessor, columns)}
     * to achieve the same effect.
     * <blockquote><pre>
     * workbook.addSheet("users", "select id, name from users where `class` = ?"
     *      , ps -&gt; ps.setString(1, "middle") ...</pre></blockquote>
     *
     * @param sql     the query SQL string
     * @param pp      the sql parameter replacement function-interface
     * @param columns the header columns
     * @return the {@link Workbook}
     * @throws SQLException if a database access error occurs
     * @deprecated use {@link #addSheet(Sheet)}
     */
    @Deprecated
    public Workbook addSheet(String sql, ParamProcessor pp, Column... columns) throws SQLException {
        return addSheet(null, sql, pp, columns);
    }

    /**
     * Add a {@link StatementSheet} to the tail with worksheet name
     * and header {@link Column} setting. The {@link ParamProcessor}
     * is a sql parameter replacement function-interface to replace "?" in
     * the sql string.
     * <p>
     * Also you can use {@code addSheet(new StatementSheet(name, connection, sql, paramProcessor, columns)}
     * to achieve the same effect.
     * <blockquote><pre>
     * workbook.addSheet("users", "select id, name from users where `class` = ?"
     *      , ps -&gt; ps.setString(1, "middle") ...</pre></blockquote>
     *
     * @param name    the worksheet name
     * @param sql     the query SQL string
     * @param pp      the sql parameter replacement function-interface
     * @param columns the header columns
     * @return the {@link Workbook}
     * @throws SQLException if a database access error occurs
     * @deprecated use {@link #addSheet(Sheet)}
     */
    @Deprecated
    public Workbook addSheet(String name, String sql, ParamProcessor pp
        , Column... columns) throws SQLException {
        PreparedStatement ps = con.prepareStatement(sql
            , ResultSet.TYPE_FORWARD_ONLY, ResultSet.CONCUR_READ_ONLY);
        return addSheet(name, ps, pp, columns);
    }

    /**
     * Add a {@link StatementSheet} to the tail with header {@link Column} setting.
     * Also you can use {@code addSheet(new StatementSheet(null, columns).setPs(ps)}
     * to achieve the same effect.
     *
     * @param ps      the {@link PreparedStatement}
     * @param columns the header columns
     * @return the {@link Workbook}
     * @throws SQLException if a database access error occurs
     * @deprecated use {@link #addSheet(Sheet)}
     */
    @Deprecated
    public Workbook addSheet(PreparedStatement ps, Column... columns) throws SQLException {
        return addSheet(null, ps, columns);
    }

    /**
     * Add a {@link StatementSheet} to the tail with worksheet name
     * and header {@link Column} setting. Also you can use
     * {@code addSheet(new StatementSheet(name, columns).setPs(ps)}
     * to achieve the same effect.
     *
     * @param name    the worksheet name
     * @param ps      the {@link PreparedStatement}
     * @param columns the header columns
     * @return the {@link Workbook}
     * @throws SQLException if a database access error occurs
     * @deprecated use {@link #addSheet(Sheet)}
     */
    @Deprecated
    public Workbook addSheet(String name, PreparedStatement ps, Column... columns) throws SQLException {
        return addSheet(name, ps, null, columns);
    }

    /**
     * Add a {@link StatementSheet} to the tail with header {@link Column} setting.
     *
     * @param ps      the {@link PreparedStatement}
     * @param pp      the sql parameter replacement function-interface
     * @param columns the header columns
     * @return the {@link Workbook}
     * @throws SQLException if a database access error occurs
     * @deprecated use {@link #addSheet(Sheet)}
     */
    @Deprecated
    public Workbook addSheet(PreparedStatement ps, ParamProcessor pp, Column... columns) throws SQLException {
        return addSheet(null, ps, pp, columns);
    }

    /**
     * Add a {@link StatementSheet} to the tail with worksheet name
     * and header {@link Column} setting.
     * <blockquote><pre>
     * workbook.addSheet("users", ps, ps -&gt; ps.setString(1, "middle") ...
     * </pre></blockquote>
     *
     * @param name    the worksheet name
     * @param ps      PreparedStatement
     * @param pp      the sql parameter replacement function-interface
     * @param columns the header columns
     * @return the {@link Workbook}
     * @throws SQLException if a database access error occurs
     * @deprecated use {@link #addSheet(Sheet)}
     */
    @Deprecated
    public Workbook addSheet(String name, PreparedStatement ps, ParamProcessor pp, Column... columns) throws SQLException {
        StatementSheet sheet = new StatementSheet(name, columns);
        try {
            ps.setFetchSize(Integer.MIN_VALUE);
            ps.setFetchDirection(ResultSet.FETCH_REVERSE);
        } catch (SQLException e) {
            what("Not support fetch size value of " + Integer.MIN_VALUE);
        }
        if (pp != null) pp.build(ps);
        sheet.setPs(ps);
        addSheet(sheet);
        return this;
    }

    /**
     * Insert a {@link Sheet} at the specified index
     *
     * @param index the index (zero-base) to insert at
     * @param sheet a worksheet
     * @return the {@link Workbook}
     */
    public Workbook insertSheet(int index, Sheet sheet) {
        ensureCapacityInternal();
        int _size = size;
        if (sheets[index] != null) {
            for (; _size > index; _size--) {
                sheets[_size] = sheets[_size - 1];
                sheets[_size].setId(sheets[_size].getId() + 1);
            }
        }
        sheets[index] = sheet;
        sheet.setId(index + 1);
        sheet.setWorkbook(this);
        size++;
        return this;
    }

    /**
     * Remove the {@link Sheet} from the specified index
     *
     * @param index the index (zero-base) to be delete
     * @return the {@link Workbook}
     */
    public Workbook remove(int index) {
        if (index < 0 || index >= size) {
            return this;
        }
        if (index == size - 1) {
            sheets[index] = null;
        } else {
            for (; index < size - 1; index++) {
                sheets[index] = sheets[index + 1];
                sheets[index].setId(sheets[index].getId() - 1);
            }
        }
        size--;
        return this;
    }

    /**
     * Returns the Sheet of the specified index
     *
     * @param index the index (zero-base)
     * @return the {@link Sheet}
     * @throws IndexOutOfBoundsException if index is negative number
     * or greater than the worksheet size in current workbook
     */
    public Sheet getSheetAt(int index) {
        if (index < 0 || index >= size)
            throw new IndexOutOfBoundsException("Index: " + index + ", Size: " + size);
        return sheets[index];
    }

    /**
     * Return a {@link Sheet} with the specified name
     * <p>
     * Note: This method can only return the {@code Sheet} which name specified when created.
     *
     * @param sheetName the sheet name
     * @return the {@link Sheet}, returns null if not found
     */
    public Sheet getSheet(String sheetName) {
        if (StringUtil.isEmpty(sheetName)) return null;
        for (Sheet sheet : sheets) {
            if (sheetName.equals(sheet.getName())) {
                return sheet;
            }
        }
        return null;
    }

    /**
     * Setting a progress watch
     *
     * <blockquote><pre>
     * new Workbook().onProgress((sheet, row) -&gt; {
     *     System.out.println(sheet + " write " + row + " rows");
     * })</pre></blockquote>
     *
     * @param progressConsumer a progress watch
     * @return the {@link Workbook}
     */
    public Workbook onProgress(BiConsumer<Sheet, Integer> progressConsumer) {
        this.progressConsumer = progressConsumer;
        return this;
    }

    /**
     * Returns progress watch
     *
     * @return progress consumer if setting
     */
    public BiConsumer<Sheet, Integer> getProgressConsumer() {
        return progressConsumer;
    }

//    /**
//     * Save as excel97~2003
//     * <p>
//     * You mast add eec-e3-support.jar into class path to support excel97~2003
//     *
//     * @return the {@link Workbook}
//     * @throws OperationNotSupportedException if eec-e3-support not import into class path
//     */
//    public Workbook saveAsExcel2003() throws OperationNotSupportedException {
//        try {
//            // Create Styles and SharedStringTable
//            Class<?> clazz = Class.forName("org.ttzero.excel.entity.e3.BIFF8WorkbookWriter");
//            Constructor<?> constructor = clazz.getDeclaredConstructor(this.getClass());
//            workbookWriter = (IWorkbookWriter) constructor.newInstance(this);
//        } catch (Exception e) {
//            throw new OperationNotSupportedException("Excel97-2003 Not support now.");
//        }
//        return this;
//    }

    /**
     * Save file as Comma-Separated Values. Each worksheet corresponds to
     * a csv file. Default charset is 'UTF8' and separator character is ','.
     *
     * @return the {@link Workbook}
     */
    public Workbook saveAsCSV() {
        workbookWriter = new CSVWorkbookWriter(this);
        return this;
    }

    private void ensureCapacityInternal() {
        if (size >= sheets.length) {
            sheets = Arrays.copyOf(sheets, size + 1);
        }
    }

    //////////////////////////Print Out/////////////////////////////

    /**
     * Export the workbook to the specified folder
     * <p>
     * If the path is a folder, save the workbook to that folder,
     * if there has duplicate file name, add '(n)' after the output
     * file name to avoid overwriting the original file.
     * If it is a file, overwrite the file.
     *
     * @param path the output pat, It can be a directory or a
     *             full file path
     * @throws IOException if I/O error occur
     */
    @Override
    public void writeTo(Path path) throws IOException {
        checkAndInitWriter();
        if (!exists(path)) {
            String name = path.getFileName().toString();
            // write to file
            if (name.indexOf('.') > 0) {
                Path parent = path.getParent();
                FileUtil.mkdir(parent);
                writeTo(path.toFile());
                return;
                // write to directory
            } else FileUtil.mkdir(path);
        } else if (!Files.isDirectory(path)) {
            writeTo(path.toFile());
            return;
        }
        workbookWriter.writeTo(path);
    }

    /**
     * Export the workbook to the specified {@link OutputStream}.
     * It mostly used for small excel file export and download
     * <blockquote><pre>
     * public void export(HttpServletResponse response) throws IOException {
     *     String fileName = java.net.URLEncoder.encode("{name}.xlsx", "UTF-8");
     *     response.setHeader(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=\"" + fileName + "\"; filename*=utf-8''" + fileName);
     *     new Workbook("{name}", "{author}")
     *     .setAutoSize(true)
     *     .addSheet(new ListSheet&lt;ListObjectSheetTest.Item&gt;("{worksheet name}", new ArrayList&lt;&gt;()))
     *     // Write to HttpServletResponse
     *     .writeTo(response.getOutputStream());
     * }</pre></blockquote>
     *
     * @param os the OutputStream
     * @throws IOException         if I/O error occur
     * @throws ExcelWriteException other runtime error
     */
    public void writeTo(OutputStream os) throws IOException, ExcelWriteException {
        checkAndInitWriter();
        workbookWriter.writeTo(os);
    }

    /**
     * Export the workbook to the specified folder
     *
     * @param file                 the output file name
     * @throws IOException         if I/O error occur
     * @throws ExcelWriteException other runtime error
     */
    public void writeTo(File file) throws IOException, ExcelWriteException {
        checkAndInitWriter();
        if (!file.getParentFile().exists()) {
            FileUtil.mkdir(file.toPath().getParent());
        }
        workbookWriter.writeTo(file);
    }

    /////////////////////////////////Template///////////////////////////////////
    private InputStream is;
    private Object o;

    /**
     * Returns the template io-stream
     *
     * @return the io-stream of template
     */
    public InputStream getTemplate() {
        return is;
    }

    /**
     * Returns the replacement object
     *
     * @return the object
     */
    public Object getBind() {
        return o;
    }

    /**
     * Bind a excel template and set an object to replace the
     * placeholder character in template
     *
     * @param is the template io-stream
     * @param o  bind a replacement object
     * @return the {@link Workbook}
     */
    public Workbook withTemplate(InputStream is, Object o) {
        this.is = is;
        this.o = o;
        return this;
    }

    protected Path template() throws IOException {
        return workbookWriter.template();
    }

    /**
     * Setting a customize workbook writer
     *
     * @param workbookWriter a customize {@link IWorkbookWriter}
     * @return the {@link Workbook}
     */
    public Workbook setWorkbookWriter(IWorkbookWriter workbookWriter) {
        this.workbookWriter = workbookWriter;
        this.workbookWriter.setWorkbook(this);
        return this;
    }

    /**
     * Create some global entry.
     */
    protected void init() {
        // Create SharedStringTable
        if (sst == null) {
            sst = new SharedStrings();
        }
        // Create a global styles
        if (styles == null) {
            styles = Styles.create(i18N);
        }
    }

    /**
     * Check and Create {@link IWorkbookWriter}
     */
    protected void checkAndInitWriter() {
        if (workbookWriter == null) {
            // Create Styles and SharedStringTable
            init();
            workbookWriter = new XMLWorkbookWriter(this);
        }
    }

    /**
     * Add a content-type
     *
     * @param type {@link ContentType.Type}
     * @return current {@link Workbook}
     */
    public Workbook addContentType(ContentType.Type type) {
        contentType.add(type);
        return this;
    }

    /**
     * Add a content-type refer
     *
     * @param rel {@link Relationship}
     * @return current {@link Workbook}
     */
    public Workbook addContentTypeRel(Relationship rel) {
        contentType.addRel(rel);
        return this;
    }

    /**
     * Returns the global ContentType
     *
     * @return {@link ContentType}
     */
    public ContentType getContentType() {
        return contentType;
    }

    /**
     * Increment and returns drawing-counter
     *
     * @return drawing-counter
     */
    public int incrementDrawingCounter() {
        return ++drawingCounter;
    }

    /**
     * Returns count of drawing object
     *
     * @return count of drawing object
     */
    public int getDrawingCounter() {
        return drawingCounter;
    }

    /**
     * Increment and returns media-counter
     *
     * @return media-counter
     */
    public int incrementMediaCounter() {
        return ++mediaCounter;
    }

    /**
     * Returns count of media object
     *
     * @return count of media object
     */
    public int getMediaCounter() {
        return mediaCounter;
    }
}

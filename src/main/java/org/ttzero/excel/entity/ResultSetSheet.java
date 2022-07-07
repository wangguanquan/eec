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

import org.ttzero.excel.reader.Cell;
import org.ttzero.excel.util.StringUtil;

import java.io.IOException;
import java.math.BigDecimal;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.sql.Timestamp;

import static java.sql.Types.BIGINT;
import static java.sql.Types.BIT;
import static java.sql.Types.CHAR;
import static java.sql.Types.DATE;
import static java.sql.Types.DECIMAL;
import static java.sql.Types.DOUBLE;
import static java.sql.Types.FLOAT;
import static java.sql.Types.INTEGER;
import static java.sql.Types.LONGVARCHAR;
import static java.sql.Types.NULL;
import static java.sql.Types.NUMERIC;
import static java.sql.Types.REAL;
import static java.sql.Types.SMALLINT;
import static java.sql.Types.TIME;
import static java.sql.Types.TIMESTAMP;
import static java.sql.Types.TINYINT;
import static java.sql.Types.VARCHAR;

/**
 * ResultSet is one of the worksheet data sources, It has a subclass
 * {@link StatementSheet}. Most of the time it is used to get the
 * result of a stored procedure.
 * <p>
 * Write data to the row-block via the cursor, finished write when
 * {@link ResultSet#next} returns false
 *
 * @see StatementSheet
 *
 * @author guanquan.wang on 2017/9/27.
 */
public class ResultSetSheet extends Sheet {
    protected ResultSet rs;

    /**
     * Constructor worksheet
     */
    public ResultSetSheet() {
        super();
    }

    /**
     * Constructor worksheet
     *
     * @param name the worksheet name
     */
    public ResultSetSheet(String name) {
        super(name);
    }

    /**
     * Constructor worksheet
     *
     * @param name    the worksheet name
     * @param columns the header info
     */
    public ResultSetSheet(String name, final org.ttzero.excel.entity.Column... columns) {
        super(name, columns);
    }

    /**
     * Constructor worksheet
     *
     * @param name      the worksheet name
     * @param waterMark the water mark
     * @param columns   the header info
     */
    public ResultSetSheet(String name, WaterMark waterMark, final org.ttzero.excel.entity.Column... columns) {
        super(name, waterMark, columns);
    }

    /**
     * Constructor worksheet
     *
     * @param rs the ResultSet
     */
    public ResultSetSheet(ResultSet rs) {
        this(null, rs);
    }

    /**
     * Constructor worksheet
     *
     * @param name the worksheet name
     * @param rs   the ResultSet
     */
    public ResultSetSheet(String name, ResultSet rs) {
        super(name);
        this.rs = rs;
    }

    /**
     * Constructor worksheet
     *
     * @param rs      the ResultSet
     * @param columns the header info
     */
    public ResultSetSheet(ResultSet rs, final org.ttzero.excel.entity.Column... columns) {
        this(null, rs, null, columns);
    }

    /**
     * Constructor worksheet
     *
     * @param name    the worksheet name
     * @param rs      the ResultSet
     * @param columns the header info
     */
    public ResultSetSheet(String name, ResultSet rs, final org.ttzero.excel.entity.Column... columns) {
        this(name, rs, null, columns);
    }

    /**
     * Constructor worksheet
     *
     * @param rs        the ResultSet
     * @param waterMark the water mark
     * @param columns   the header info
     */
    public ResultSetSheet(ResultSet rs, WaterMark waterMark, final org.ttzero.excel.entity.Column... columns) {
        this(null, rs, waterMark, columns);
    }

    /**
     * Constructor worksheet
     *
     * @param name      the worksheet name
     * @param rs        the ResultSet
     * @param waterMark the water mark
     * @param columns   the header info
     */
    public ResultSetSheet(String name, ResultSet rs, WaterMark waterMark, final org.ttzero.excel.entity.Column... columns) {
        super(name, waterMark, columns);
        this.rs = rs;
    }

    /**
     * Setting ResultSet
     *
     * @param rs the ResultSet
     * @return {@code ResultSetSheet}
     */
    public ResultSetSheet setRs(ResultSet rs) {
        this.rs = rs;
        return this;
    }

    /**
     * Release resources
     *
     * @throws IOException if I/O error occur
     */
    @Override
    public void close() throws IOException {
        if (shouldClose && rs != null) {
            try {
                rs.close();
            } catch (SQLException e) {
                workbook.what("9006", e.getMessage());
            }
        }
        super.close();
    }

    /**
     * Reset the row-block data
     */
    @Override
    protected void resetBlockData() {
        int len = columns.length, n = 0, limit = getRowLimit();

        try {
            for (int rbs = getRowBlockSize(); n++ < rbs && rows < limit && rs.next(); rows++) {
                Row row = rowBlock.next();
                row.index = rows;
                Cell[] cells = row.realloc(len);
                for (int i = 1; i <= len; i++) {
                    SQLColumn hc = (SQLColumn) columns[i - 1];

                    // clear cells
                    Cell cell = cells[i - 1];
                    cell.clear();

                    Object e;
                    switch (hc.sqlType) {
                        case VARCHAR:
                        case LONGVARCHAR:
                        case NULL:        e = rs.getString(i);     break;
                        case INTEGER:
                        case TINYINT:
                        case SMALLINT:
                        case BIT:
                        case CHAR:        e = rs.getInt(i);        break;
                        case DATE:        e = rs.getDate(i);       break;
                        case TIMESTAMP:   e = rs.getTimestamp(i);  break;
                        case NUMERIC:
                        case DECIMAL:     e = rs.getBigDecimal(i); break;
                        case BIGINT:      e = rs.getLong(i);       break;
                        case REAL:
                        case FLOAT:
                        case DOUBLE:      e = rs.getDouble(i);     break;
                        case TIME:        e = rs.getTime(i);       break;
                        default:          e = rs.getObject(i);     break;
                    }

                    cellValueAndStyle.reset(rows, cell, e, hc);
                    //setStyleDesign(rs,cell,hc);
                }
            }
        } catch (SQLException e) {
            throw new ExcelWriteException(e);
        }

        // Paging
        if (rows >= limit) {
            shouldClose = false;
            ResultSetSheet copy = getClass().cast(clone());
            workbook.insertSheet(id, copy);
        } else shouldClose = true;
    }

    /**
     * Get header information, get from MetaData if not specified
     * The copy sheet will use the parent worksheet header information.
     *
     * @return the header information
     */
    @Override
    protected org.ttzero.excel.entity.Column[] getHeaderColumns() {
        if (headerReady) return columns;
        if (rs == null) {
            throw new ExcelWriteException("Constructor worksheet error.\nMiss the parameter ResultSet");
        }
        int i = 0;
        try {
            ResultSetMetaData metaData = rs.getMetaData();
            if (hasHeaderColumns()) {
                org.ttzero.excel.entity.Column[] newColumns = new SQLColumn[columns.length];
                for (; i < columns.length; i++) {
                    SQLColumn column = SQLColumn.of(columns[i]);
                    newColumns[i] = column;
                    if (column.tail != null) {
                        column = (SQLColumn) column.tail;
                    }
                    if (StringUtil.isEmpty(column.getName())) {
                        column.setName(metaData.getColumnLabel(i + 1));
                    }
                    // FIXME maybe do not reset the types
                    column.sqlType = metaData.getColumnType(i + 1);
                    Class<?> metaClazz = columnTypeToClass(column.sqlType);
                    if (column.clazz != metaClazz) {
                        what("The specified type " + column.clazz + " is different" +
                            " from metadata column type " + metaClazz);
                        column.clazz = metaClazz;
                    }
                }
                columns = newColumns;
            } else {
                int count = metaData.getColumnCount();
                columns = new org.ttzero.excel.entity.Column[count];
                for (; ++i <= count; ) {
                    columns[i - 1] = new SQLColumn(metaData.getColumnLabel(i), metaData.getColumnType(i)
                        , columnTypeToClass(metaData.getColumnType(i)));
                }
            }
        } catch (SQLException e) {
            what("un-support get result set meta data.");
        }

        if (hasHeaderColumns()) {

            for (i = 0; i < columns.length; i++) {
                if (StringUtil.isEmpty(columns[i].getName())) {
                    columns[i].setName(String.valueOf(i));
                }
            }
        }
        return columns;
    }

    /**
     * Convert {@link java.sql.Types} to java type
     *
     * @param type type sql type
     * @return java class
     */
    protected Class<?> columnTypeToClass(int type) {
        Class<?> clazz;
        switch (type) {
            case VARCHAR:
            case CHAR:
            case LONGVARCHAR:
            case NULL:      clazz = String.class;        break;
            case INTEGER:   clazz = Integer.class;       break;
            case DATE:      clazz = java.sql.Date.class; break;
            case TIMESTAMP: clazz = Timestamp.class;     break;
            case NUMERIC:
            case DECIMAL:   clazz = BigDecimal.class;    break;
            case BIT:       clazz = Boolean.class;       break;
            case TINYINT:   clazz = Byte.class;          break;
            case SMALLINT:  clazz = Short.class;         break;
            case BIGINT:    clazz = Long.class;          break;
            case REAL:      clazz = Float.class;         break;
            case FLOAT:
            case DOUBLE:    clazz = Double.class;        break;
//            case CHAR:      clazz = char.class;          break;
            case TIME:      clazz = java.sql.Time.class; break;
            default:        clazz = Object.class;
        }
        return clazz;
    }

    private static class SQLColumn extends org.ttzero.excel.entity.Column {
        int sqlType;

        public SQLColumn(String name, int sqlType, Class<?> clazz) {
            super(name, clazz);
            this.sqlType = sqlType;
        }

        public SQLColumn(org.ttzero.excel.entity.Column other) {
            this.key = other.key;
            this.name = other.name;
            this.clazz = other.clazz;
            this.share = other.share;
            this.processor = other.processor;
            this.styleProcessor = other.styleProcessor;
            this.cellStyle = other.cellStyle;
            this.width = other.width;
            this.o = other.o;
            this.styles = other.styles;
            this.headerComment = other.headerComment;
            this.cellComment = other.cellComment;
            this.numFmt = other.numFmt;
            this.ignoreValue = other.ignoreValue;
            this.wrapText = other.wrapText;
            if (other.next != null) {
                addSubColumn(new SQLColumn(other.next));
            }
        }

        public static SQLColumn of(org.ttzero.excel.entity.Column other) {
            return new SQLColumn(other);
        }
    }
}

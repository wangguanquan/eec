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

package cn.ttzero.excel.entity;

import cn.ttzero.excel.reader.Cell;
import cn.ttzero.excel.util.StringUtil;

import java.io.IOException;
import java.math.BigDecimal;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.sql.Timestamp;

import static java.sql.Types.*;
import static java.sql.Types.TIME;

/**
 * ResultSet is one of the worksheet data sources, It has a subclass
 * {@code StatementSheet}. Most of the time it is used to get the
 * result of a stored procedure.
 * <p>
 * Write data to the row-block via the cursor, finished write when
 * {@code ResultSet#next} returns false
 * <p>
 * Created by guanquan.wang on 2017/9/27.
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
    public ResultSetSheet(String name, final Column... columns) {
        super(name, columns);
    }

    /**
     * Constructor worksheet
     *
     * @param name      the worksheet name
     * @param waterMark the water mark
     * @param columns   the header info
     */
    public ResultSetSheet(String name, WaterMark waterMark, final Column... columns) {
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
    public ResultSetSheet(ResultSet rs, final Column... columns) {
        this(null, rs, null, columns);
    }

    /**
     * Constructor worksheet
     *
     * @param name    the worksheet name
     * @param rs      the ResultSet
     * @param columns the header info
     */
    public ResultSetSheet(String name, ResultSet rs, final Column... columns) {
        this(name, rs, null, columns);
    }

    /**
     * Constructor worksheet
     *
     * @param rs        the ResultSet
     * @param waterMark the water mark
     * @param columns   the header info
     */
    public ResultSetSheet(ResultSet rs, WaterMark waterMark, final Column... columns) {
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
    public ResultSetSheet(String name, ResultSet rs, WaterMark waterMark, final Column... columns) {
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
     * @throws IOException if io error occur
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
        int len = columns.length, n = 0, limit = sheetWriter.getRowLimit() - 1;

        try {
            for (int rbs = getRowBlockSize(); n++ < rbs && rows < limit && rs.next(); rows++) {
                Row row = rowBlock.next();
                row.index = rows;
                Cell[] cells = row.realloc(len);
                for (int i = 1; i <= len; i++) {
                    Column hc = columns[i - 1];

                    // clear cells
                    Cell cell = cells[i - 1];
                    cell.clear();

                    Object e = rs.getObject(i);

                    // blank cell
                    if (e == null) {
                        cell.setBlank();
                        continue;
                    }

                    setCellValueAndStyle(cell, e, hc);
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
    public Column[] getHeaderColumns() {
        if (headerReady) return columns;
        if (rs == null) {
            throw new ExcelWriteException("Constructor worksheet error.\nMiss the parameter ResultSet");
        }
        int i = 0;
        try {
            ResultSetMetaData metaData = rs.getMetaData();
            if (hasHeaderColumns()) {
                for (; i < columns.length; i++) {
                    Column column = columns[i];
                    if (StringUtil.isEmpty(column.getName())) {
                        column.setName(metaData.getColumnName(i + 1));
                    }
                    // FIXME maybe do not reset the types
                    Class<?> metaClazz = columnTypeToClass(metaData.getColumnType(i + 1));
                    if (column.clazz != metaClazz) {
                        what("The specified type " + column.clazz + " is different" +
                            " from metadata column type " + metaClazz);
                        column.clazz = metaClazz;
                    }
                }
            } else {
                int count = metaData.getColumnCount();
                columns = new Column[count];
                for (; ++i <= count; ) {
                    columns[i - 1] = new Column(metaData.getColumnName(i)
                        , columnTypeToClass(metaData.getColumnType(i)));
                }
            }
        } catch (SQLException e) {
            what("un-support get result set meta data.");
        }

        if (hasHeaderColumns()) {
            // Check the limit of columns
            checkColumnLimit();

            for (i = 0; i < columns.length; i++) {
                if (StringUtil.isEmpty(columns[i].getName())) {
                    columns[i].setName(String.valueOf(i));
                }
                if (columns[i].styles == null) {
                    columns[i].styles = workbook.getStyles();
                }
            }
            headerReady = columns.length > 0;
        }
        return columns;
    }

    /**
     * Convert {@code java.sql.Type} to java type
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
            case INTEGER:   clazz = int.class;           break;
            case DATE:      clazz = java.sql.Date.class; break;
            case TIMESTAMP: clazz = Timestamp.class;     break;
            case NUMERIC:
            case DECIMAL:   clazz = BigDecimal.class;    break;
            case BIT:       clazz = boolean.class;       break;
            case TINYINT:   clazz = byte.class;          break;
            case SMALLINT:  clazz = short.class;         break;
            case BIGINT:    clazz = long.class;          break;
            case REAL:      clazz = float.class;         break;
            case FLOAT:
            case DOUBLE:    clazz = double.class;        break;
//            case CHAR:      clazz = char.class;          break;
            case TIME:      clazz = java.sql.Time.class; break;
            default:        clazz = Object.class;
        }
        return clazz;
    }
}

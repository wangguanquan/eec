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

import org.ttzero.excel.manager.Const;
import org.ttzero.excel.processor.StyleProcessor;
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
 * {@code ResultSetSheet}的数据源为{@link ResultSet}一般情况下它用于存储过程，
 * {@code ResultSetSheet}可以将存储过程的查询结果直接转为工作表的数据，省掉将查结果转为
 * Java对象再转为工作表输出协议的数据结构。
 *
 * <p>如果未指定{@code Columns}表头时将从{@link ResultSetMetaData}源数据中获取，
 * 优先使用别名做为表头，列顺序与query字段一致</p>
 *
 * <p>这是一个比较小众的工作表，最好只在比较简单的场景下使用，比如一次性导出的场景。
 * 因为{@code StatementSheet}并不支持数据切片，所以当查询结果较大时可能出现OOM。
 * 如果不确认数据量时最好使用{@link ListSheet}分片获取数据</p>
 *
 * @author guanquan.wang on 2017/9/27.
 * @see StatementSheet
 */
public class ResultSetSheet extends Sheet {
    /**
     * 数据源ResultSet
     */
    protected ResultSet rs;
    /**
     * 行级动态样式处理器
     */
    private StyleProcessor<ResultSet> styleProcessor;

    /**
     * 实例化工作表，未指定工作表名称时默认以{@code 'Sheet'+id}命名
     */
    public ResultSetSheet() {
        super();
    }

    /**
     * 实例化工作表并指定工作表名称
     *
     * @param name 工作表名称
     */
    public ResultSetSheet(String name) {
        super(name);
    }

    /**
     * 实例化工作表并指定表头信息
     *
     * @param columns 表头信息
     */
    public ResultSetSheet(final Column... columns) {
        super(columns);
    }

    /**
     * 实例化工作表并指定工作表名称和表头信息
     *
     * @param name    工作表名称
     * @param columns 表头信息
     */
    public ResultSetSheet(String name, final Column... columns) {
        super(name, columns);
    }

    /**
     * 实例化工作表并指定工作表名称，水印和表头信息
     *
     * @param name      工作表名称
     * @param waterMark 水印
     * @param columns   表头信息
     */
    public ResultSetSheet(String name, WaterMark waterMark, final Column... columns) {
        super(name, waterMark, columns);
    }

    /**
     * 实例化工作表并指定数据源{@code ResultSet}
     *
     * @param rs 数据源{@code ResultSet}
     */
    public ResultSetSheet(ResultSet rs) {
        this(null, rs);
    }

    /**
     * 实例化工作表并指定工作表名和数据源{@code ResultSet}
     *
     * @param name 工作表名
     * @param rs   数据源{@code ResultSet}
     */
    public ResultSetSheet(String name, ResultSet rs) {
        super(name);
        this.rs = rs;
    }

    /**
     * 实例化工作表并指定数据源{@code ResultSet}和表头信息
     *
     * @param rs      数据源{@code ResultSet}
     * @param columns 表头信息
     */
    public ResultSetSheet(ResultSet rs, final Column... columns) {
        this(null, rs, null, columns);
    }

    /**
     * 实例化工作表并指定工作表名、数据源{@code ResultSet}和表头信息
     *
     * @param name    工作表名
     * @param rs      数据源{@code ResultSet}
     * @param columns 表头信息
     */
    public ResultSetSheet(String name, ResultSet rs, final Column... columns) {
        this(name, rs, null, columns);
    }

    /**
     * 实例化工作表并指定数据源{@code ResultSet}、水印和表头信息
     *
     * @param rs        数据源{@code ResultSet}
     * @param waterMark 水印
     * @param columns   表头信息
     */
    public ResultSetSheet(ResultSet rs, WaterMark waterMark, final Column... columns) {
        this(null, rs, waterMark, columns);
    }

    /**
     * 实例化工作表并指定工作表名、数据源{@code ResultSet}、水印和表头信息
     *
     * @param name      工作表名
     * @param rs        数据源{@code ResultSet}
     * @param waterMark 水印
     * @param columns   表头信息
     */
    public ResultSetSheet(String name, ResultSet rs, WaterMark waterMark, final Column... columns) {
        super(name, waterMark, columns);
        this.rs = rs;
    }

    /**
     * 设置数据源{@code ResultSet}
     *
     * @param rs 数据源{@code ResultSet}
     * @return 当前工作表
     * @deprecated 使用 {@link #setResultSet(ResultSet)}替换
     */
    @Deprecated
    public ResultSetSheet setRs(ResultSet rs) {
        return setResultSet(rs);
    }

    /**
     * 设置数据源{@code ResultSet}
     *
     * @param resultSet 数据源{@code ResultSet}
     * @return 当前工作表
     */
    public ResultSetSheet setResultSet(ResultSet resultSet) {
        this.rs = resultSet;
        return this;
    }

    /**
     * 设置行级动态样式处理器，作用于整行优先级高于单元格动态样式处理器
     *
     * @param styleProcessor 行级动态样式处理器
     * @return 当前工作表
     */
    public Sheet setStyleProcessor(StyleProcessor<ResultSet> styleProcessor) {
        this.styleProcessor = styleProcessor;
        putExtProp(Const.ExtendPropertyKey.STYLE_DESIGN, styleProcessor);
        return this;
    }

    /**
     * 获取当前工作表的行级动态样式处理器，如果未设置则从扩展参数中查找
     *
     * @return 行级动态样式处理器
     */
    public StyleProcessor<ResultSet> getStyleProcessor() {
        if (styleProcessor != null) return styleProcessor;
        @SuppressWarnings("unchecked")
        StyleProcessor<ResultSet> fromExtProp = (StyleProcessor<ResultSet>) getExtPropValue(Const.ExtendPropertyKey.STYLE_DESIGN);
        return this.styleProcessor = fromExtProp;
    }

    /**
     * 关闭数据源并关闭{@code ResultSet}
     *
     * @throws IOException if I/O error occur
     */
    @Override
    public void close() throws IOException {
        if (shouldClose && rs != null) {
            try {
                rs.close();
            } catch (SQLException e) {
                LOGGER.warn("Close ResultSet error.", e);
            }
        }
        super.close();
    }

    /**
     * 重置{@code RowBlock}行块数据
     */
    @Override
    protected void resetBlockData() {
        int len = columns.length, n = 0, limit = getRowLimit();
        boolean hasGlobalStyleProcessor = (extPropMark & 2) == 2;
        try {
            for (int rbs = rowBlock.capacity(); n++ < rbs && rows < limit && rs.next(); rows++) {
                Row row = rowBlock.next();
                row.index = rows;
                row.height = getRowHeight();
                Cell[] cells = row.realloc(len);
                for (int i = 1; i <= len; i++) {
                    SQLColumn hc = (SQLColumn) columns[i - 1];

                    // clear cells
                    Cell cell = cells[i - 1];
                    cell.clear();

                    Object e;
                    if (hc.ri > 0) {
                        switch (hc.sqlType) {
                            case VARCHAR:
                            case LONGVARCHAR:
                            case NULL:           e = rs.getString(hc.ri);      break;
                            case INTEGER:
                            case TINYINT:
                            case SMALLINT:
                            case BIT:
                            case CHAR:            e = rs.getInt(hc.ri);        break;
                            case DATE:            e = rs.getDate(hc.ri);       break;
                            case TIMESTAMP:       e = rs.getTimestamp(hc.ri);  break;
                            case NUMERIC:
                            case DECIMAL:         e = rs.getBigDecimal(hc.ri); break;
                            case BIGINT:          e = rs.getLong(hc.ri);       break;
                            case REAL:
                            case FLOAT:
                            case DOUBLE:          e = rs.getDouble(hc.ri);     break;
                            case TIME:            e = rs.getTime(hc.ri);       break;
                            default:              e = rs.getObject(hc.ri);     break;
                        }
                    } else e = null;

                    cellValueAndStyle.reset(row, cell, e, hc);
                    if (hasGlobalStyleProcessor) {
                        cellValueAndStyle.setStyleDesign(rs, cell, hc, getStyleProcessor());
                    }
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
     * 获取表头，未指定表头时从{@link ResultSetMetaData}源数据中获取，
     * 优先使用别名做为表头，列顺序与query字段一致
     *
     * @return 表头信息
     */
    @Override
    protected Column[] getHeaderColumns() {
        if (headerReady) return columns;
        if (rs == null) {
            throw new ExcelWriteException("Constructor worksheet error.\nMiss the parameter ResultSet");
        }
        int i = 0;
        try {
            ResultSetMetaData metaData = rs.getMetaData();
            int count = metaData.getColumnCount();
            if (hasHeaderColumns()) {
                Column[] newColumns = new SQLColumn[columns.length];
                for (; i < columns.length; i++) {
                    SQLColumn column = SQLColumn.of(columns[i]);
                    newColumns[i] = column;
                    if (column.tail != null) column = (SQLColumn) column.tail;
                    if (i + 1 > count) {
                        LOGGER.warn("Column [{}] cannot be mapped.", columns[i].getName());
                        continue;
                    }
                    if (StringUtil.isEmpty(column.getName()))
                        column.setName(metaData.getColumnLabel(i + 1));
                    column.ri = StringUtil.isNotEmpty(column.key) ? findByKey(metaData, column.key) : i + 1;

                    if (column.ri < 0) {
                        LOGGER.warn("Column [{}] cannot be mapped.", columns[i].getName());
                        continue;
                    }

                    column.sqlType = metaData.getColumnType(i + 1);
                    Class<?> metaClazz = columnTypeToClass(column.sqlType);
                    if (column.clazz != metaClazz) {
                        LOGGER.warn("The specified type {} is different from metadata column type {}", column.clazz, metaClazz);
//                        column.clazz = metaClazz;
                    }
                }
                columns = newColumns;
            } else {
                columns = new Column[count];
                while (++i <= count) {
                    SQLColumn column = new SQLColumn(metaData.getColumnLabel(i), metaData.getColumnType(i)
                        , columnTypeToClass(metaData.getColumnType(i)));
                    column.ri = StringUtil.isNotEmpty(column.key) ? findByKey(metaData, column.key) : i;
                    columns[i - 1] = column;
                }
            }
        } catch (SQLException e) {
            LOGGER.warn("Get meta data occur error.", e);
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

    protected int findByKey(ResultSetMetaData metaData, String key) throws SQLException {
        for (int i = 1, len = metaData.getColumnCount(); i <= len; i++) {
            if (key.equals(metaData.getColumnLabel(i))) {
                return i;
            }
        }
        return -1;
    }

    /**
     * 将SQL类型{@link java.sql.Types}转换为Java类型
     *
     * @param type SQL类型{@code java.sql.Types}
     * @return Java类型
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

    /**
     * {@code ResultSetSheet}独有的列对象，除了{@link Column}包含的信息外，它还保存当列对应的SQL类型和
     * {@code ResultSet}下标，有了下标后续列取值可直接根据{@code ri}直接取值
     */
    public static class SQLColumn extends Column {
        /**
         * SQL类型，等同于{@link java.sql.Types}中的静态类型
         */
        public int sqlType;
        /**
         * ResultSet下标
         */
        public int ri;

        public SQLColumn(String name, int sqlType, Class<?> clazz) {
            super(name, clazz);
            this.sqlType = sqlType;
        }

        public SQLColumn(Column other) {
            super.from(other);
            if (other instanceof SQLColumn) {
                SQLColumn o = (SQLColumn) other;
                this.sqlType = o.sqlType;
                this.ri = o.ri;
            }
            if (other.next != null) {
                addSubColumn(new SQLColumn(other.next));
            }
        }

        public static SQLColumn of(Column other) {
            return new SQLColumn(other);
        }
    }
}

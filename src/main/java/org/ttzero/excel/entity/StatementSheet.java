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

import org.ttzero.excel.processor.ParamProcessor;

import java.io.IOException;
import java.nio.file.Path;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;

/**
 * Statement is one of the worksheet data sources, it's
 * extends from {@link ResultSetSheet}, and will be obtained from
 * MetaData if no header information is given. When MetaData
 * cannot be obtained, the header name will be setting as 1, 2,
 * and 3...
 * <p>
 * The Connection will not be actively closed, but the {@link java.sql.Statement}
 * and {@link ResultSet} will be closed with worksheet.
 *
 * @see ResultSetSheet
 *
 * @author guanquan.wang on 2017/9/26.
 */
public class StatementSheet extends ResultSetSheet {
    private PreparedStatement ps;

    /**
     * Constructor worksheet
     */
    public StatementSheet() {
        super();
    }

    /**
     * Constructor worksheet
     *
     * @param name the worksheet name
     */
    public StatementSheet(String name) {
        super(name);
    }

    /**
     * Constructor worksheet
     *
     * @param columns the header info
     */
    public StatementSheet(final Column... columns) {
        super(columns);
    }

    /**
     * Constructor worksheet
     *
     * @param name    the worksheet name
     * @param columns the header info
     */
    public StatementSheet(String name, final Column... columns) {
        super(name, columns);
    }

    /**
     * Constructor worksheet
     *
     * @param name      the worksheet name
     * @param waterMark the water mark
     * @param columns   the header info
     */
    public StatementSheet(String name, WaterMark waterMark, final Column... columns) {
        super(name, waterMark, columns);
    }

    /**
     * Constructor worksheet
     *
     * @param con the Connection
     * @param sql the query SQL string
     */
    public StatementSheet(Connection con, String sql) {
        this(null, con, sql);
    }

    /**
     * Constructor worksheet
     *
     * @param name the worksheet name
     * @param con  the Connection
     * @param sql  the query SQL string
     */
    public StatementSheet(String name, Connection con, String sql) {
        super(name);
        PreparedStatement ps = null;
        try {
            ps = con.prepareStatement(sql, ResultSet.TYPE_FORWARD_ONLY, ResultSet.CONCUR_READ_ONLY);
            ps.setFetchSize(Integer.MIN_VALUE);
            ps.setFetchDirection(ResultSet.FETCH_REVERSE);
        } catch (SQLException e) {
            LOGGER.debug("Not support fetch size value of {}", Integer.MIN_VALUE);
        }
        if (ps == null) {
            throw new ExcelWriteException("Constructor worksheet error.\nMiss the parameter Statement");
        }
        this.ps = ps;
    }

    /**
     * Constructor worksheet
     *
     * @param con the Connection
     * @param sql the query SQL string
     * @param pp  the sql parameter replacement function-interface
     */
    public StatementSheet(Connection con, String sql, ParamProcessor pp) {
        this(null, con, sql, pp);
    }

    /**
     * Constructor worksheet
     *
     * @param name the worksheet name
     * @param con  the Connection
     * @param sql  the query SQL string
     * @param pp   the sql parameter replacement function-interface
     */
    public StatementSheet(String name, Connection con, String sql, ParamProcessor pp) {
        super(name);
        PreparedStatement ps = null;
        try {
            ps = con.prepareStatement(sql, ResultSet.TYPE_FORWARD_ONLY, ResultSet.CONCUR_READ_ONLY);
            ps.setFetchSize(Integer.MIN_VALUE);
            ps.setFetchDirection(ResultSet.FETCH_REVERSE);
        } catch (SQLException e) {
            LOGGER.debug("Not support fetch size value of {}", Integer.MIN_VALUE);
        }
        if (ps == null) {
            throw new ExcelWriteException("Constructor worksheet error.\nMiss the parameter Statement");
        }
        if (pp != null) {
            try {
                pp.build(ps);
            } catch (SQLException e) {
                throw new ExcelWriteException(e);
            }
        }
        this.ps = ps;
    }

    /**
     * Constructor worksheet
     *
     * @param con     the Connection
     * @param sql     the sql string
     * @param columns the header column
     */
    public StatementSheet(Connection con, String sql, Column... columns) {
        this(null, con, sql, columns);
    }

    /**
     * Constructor worksheet
     *
     * @param name    the worksheet name
     * @param con     the Connection
     * @param sql     the sql string
     * @param columns the header column
     */
    public StatementSheet(String name, Connection con, String sql, Column... columns) {
        this(name, con, sql, null, columns);
    }

    /**
     * Constructor worksheet
     *
     * @param con     the Connection
     * @param sql     the sql string
     * @param pp      the sql parameter replacement function-interface
     * @param columns the header column
     */
    public StatementSheet(Connection con, String sql, ParamProcessor pp, Column... columns) {
        this(null, con, sql, pp, columns);
    }

    /**
     * Constructor worksheet
     *
     * @param name    the worksheet name
     * @param con     the Connection
     * @param sql     the sql string
     * @param pp      the sql parameter replacement function-interface
     * @param columns the header column
     */
    public StatementSheet(String name, Connection con, String sql, ParamProcessor pp, Column... columns) {
        super(name, columns);
        PreparedStatement ps = null;
        try {
            ps = con.prepareStatement(sql, ResultSet.TYPE_FORWARD_ONLY, ResultSet.CONCUR_READ_ONLY);
            ps.setFetchSize(Integer.MIN_VALUE);
            ps.setFetchDirection(ResultSet.FETCH_REVERSE);
        } catch (SQLException e) {
            LOGGER.debug("Not support fetch size value of {}", Integer.MIN_VALUE);
        }
        if (ps == null) {
            throw new ExcelWriteException("Constructor worksheet error.\nMiss the parameter Statement");
        }
        if (pp != null) {
            try {
                pp.build(ps);
            } catch (SQLException e) {
                throw new ExcelWriteException(e);
            }
        }
        this.ps = ps;
    }

    /**
     * Setting PreparedStatement
     *
     * @param ps PreparedStatement
     * @return the {@link StatementSheet}
     * @deprecated replace with {@link #setStatement(PreparedStatement)}
     */
    @Deprecated
    public StatementSheet setPs(PreparedStatement ps) {
        return setStatement(ps);
    }

    /**
     * Setting PreparedStatement
     *
     * @param ps PreparedStatement
     * @return the {@link StatementSheet}
     */
    public StatementSheet setStatement(PreparedStatement ps) {
        this.ps = ps;
        return this;
    }

    /**
     * Release resources
     * will close the {@code Statement} and {@code ResultSet}
     *
     * @throws IOException if I/O error occur
     */
    @Override
    public void close() throws IOException {
        super.close();
        if (shouldClose && ps != null) {
            try {
                ps.close();
            } catch (SQLException e) {
                LOGGER.warn("Close ResultSet error.", e);
            }
        }
    }

    /**
     * write worksheet data to path
     *
     * @param path the storage path
     * @throws IOException if I/O error occur
     */
    @Override
    public void writeTo(Path path) throws IOException {
        if (sheetWriter != null) {
            if (!copySheet) {
                if (ps == null) {
                    throw new ExcelWriteException("Constructor worksheet error.\nMiss the parameter Statement");
                }
                // Execute query
                try {
                    rs = ps.executeQuery();
                } catch (SQLException e) {
                    throw new ExcelWriteException(e);
                }

                // Check the header information is exists
                getAndSortHeaderColumns();
            }

            if (rowBlock == null) {
                rowBlock = new RowBlock(getRowBlockSize());
            } else rowBlock.reopen();

            sheetWriter.writeTo(path);
        } else {
            throw new ExcelWriteException("Worksheet writer is not instanced.");
        }
    }
}

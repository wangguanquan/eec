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
 * {@code StatementSheet}的数据源为{@link ResultSet}，它也是{@code ResultSetSheet}的子类，
 * 用于将数据库的数据导出到Excel，它并不限制数据库类型，只需实现jdbc协议即可，使用{@code StatementSheet}
 * 可以避免将查询结果转为Java实体。
 *
 * <p>这是一个比较小众的工作表，最好只在比较简单的场景下使用，比如一次性导出的场景。因为{@code StatementSheet}
 * 并不支持数据切片，所以当查询结果较大时可能出现OOM。如果不确认数据量时最好使用{@link ListSheet}分片获取数据</p>
 *
 * <blockquote><pre>
 * try (Connection con = getConnection()) {
 *     String sql = "select name,age,create_date,update_date " +
 *               "from student where id between ? and ?";
 *     new Workbook()
 *         .addSheet(new StatementSheet(con, sql, ps -> {
 *             ps.setInt(1, 10);
 *             ps.setInt(2, 20);
 *         }))
 *         .writeTo(Paths.get("/tmp/student.xlsx"));
 * }</pre></blockquote>
 *
 * @author guanquan.wang on 2017/9/26.
 * @see ResultSetSheet
 */
public class StatementSheet extends ResultSetSheet {
    private PreparedStatement ps;

    /**
     * 实例化工作表，未指定工作表名称时默认以{@code 'Sheet'+id}命名
     */
    public StatementSheet() {
        super();
    }

    /**
     * 实例化工作表并指定工作表名称
     *
     * @param name 工作表名称
     */
    public StatementSheet(String name) {
        super(name);
    }

    /**
     * 实例化工作表并指定表头信息
     *
     * @param columns 表头信息
     */
    public StatementSheet(final Column... columns) {
        super(columns);
    }

    /**
     * 实例化工作表并指定工作表名称和表头信息
     *
     * @param name    工作表名称
     * @param columns 表头信息
     */
    public StatementSheet(String name, final Column... columns) {
        super(name, columns);
    }

    /**
     * 实例化工作表并指定工作表名称，水印和表头信息
     *
     * @param name      工作表名称
     * @param waterMark 水印
     * @param columns   表头信息
     */
    public StatementSheet(String name, WaterMark waterMark, final Column... columns) {
        super(name, waterMark, columns);
    }

    /**
     * 实例化工作表
     *
     * @param con 数据库连接 {@code Connection}
     * @param sql SQL语句
     */
    public StatementSheet(Connection con, String sql) {
        this(null, con, sql);
    }

    /**
     * 实例化工作表并指定工作表名
     *
     * @param name 工作表名
     * @param con  数据库连接 {@code Connection}
     * @param sql  SQL语句
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
     * 实例化工作表并指定工作表名
     *
     * @param con 数据库连接 {@code Connection}
     * @param sql SQL语句
     * @param pp  参数处理器
     */
    public StatementSheet(Connection con, String sql, ParamProcessor pp) {
        this(null, con, sql, pp);
    }

    /**
     * 实例化工作表并指定工作表名
     *
     * @param name 工作表名
     * @param con  数据库连接 {@code Connection}
     * @param sql  SQL语句
     * @param pp   参数处理器
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
     * 实例化工作表并指定工作表名
     *
     * @param con     数据库连接 {@code Connection}
     * @param sql     SQL语句
     * @param columns 表头信息
     */
    public StatementSheet(Connection con, String sql, Column... columns) {
        this(null, con, sql, columns);
    }

    /**
     * 实例化工作表并指定工作表名
     *
     * @param name    工作表名
     * @param con     数据库连接 {@code Connection}
     * @param sql     SQL语句
     * @param columns 表头信息
     */
    public StatementSheet(String name, Connection con, String sql, Column... columns) {
        this(name, con, sql, null, columns);
    }

    /**
     * 实例化工作表并指定工作表名
     *
     * @param con     数据库连接 {@code Connection}
     * @param sql     SQL语句
     * @param pp      参数处理器
     * @param columns 表头信息
     */
    public StatementSheet(Connection con, String sql, ParamProcessor pp, Column... columns) {
        this(null, con, sql, pp, columns);
    }

    /**
     * 实例化工作表并指定工作表名
     *
     * @param name    工作表名
     * @param con     数据库连接 {@code Connection}
     * @param sql     SQL语句
     * @param pp      参数处理器
     * @param columns 表头信息
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
     * 设置数据源{@code PreparedStatement}
     *
     * @param ps 数据源{@code PreparedStatement}
     * @return 当前工作表
     * @deprecated 使用 {@link #setStatement(PreparedStatement)}替代
     */
    @Deprecated
    public StatementSheet setPs(PreparedStatement ps) {
        return setStatement(ps);
    }

    /**
     * 设置数据源{@code PreparedStatement}
     *
     * @param ps 数据源{@code PreparedStatement}
     * @return 当前工作表
     */
    public StatementSheet setStatement(PreparedStatement ps) {
        this.ps = ps;
        return this;
    }

    /**
     * 关闭数据源并关闭{@code Statement} and {@code ResultSet}
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
     * 落盘，将工作表写到指定路径
     *
     * @param path 指定保存路径
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

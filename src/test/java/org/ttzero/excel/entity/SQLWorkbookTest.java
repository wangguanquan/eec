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


package org.ttzero.excel.entity;

import org.junit.Before;
import org.ttzero.excel.util.StringUtil;

import java.io.IOException;
import java.sql.Connection;
import java.sql.Date;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Timestamp;
import java.sql.Types;
import java.util.Properties;

/**
 * @author guanquan.wang at 2019-04-28 21:50
 */
public class SQLWorkbookTest extends WorkbookTest {
    private static final Properties pro;
    protected static String protocol;
    static {
        pro = new Properties();
        try {
            pro.load(SQLWorkbookTest.class.getClassLoader().getResourceAsStream("test.properties"));
        } catch (IOException e) {
            e.printStackTrace();
            System.exit(-1);
        }
        try {
            Class.forName(pro.getProperty("dataSource.driverClassName"));
        } catch (ClassNotFoundException e) {
            e.printStackTrace();
            System.exit(-1);
        }

        String url = pro.getProperty("dataSource.url");
        if (StringUtil.isNotEmpty(url)) {
            int i1 = url.indexOf(':'), i2 = url.indexOf(':', ++i1);
            if (i2 > i1 && i1 > 0) {
                protocol = url.substring(i1, i2);
            }
        }
        if (StringUtil.isEmpty(protocol)) {
            throw new IllegalArgumentException("dataSource.url");
        }
    }

    protected Connection getConnection() throws SQLException {
        return DriverManager.getConnection(pro.getProperty("dataSource.url")
            , pro.getProperty("dataSource.username"), pro.getProperty("dataSource.password"));
    }

    /**
     * Install test data
     */
    @Before public void init() {
        Connection con = null;
        PreparedStatement ps = null;
        ResultSet rs = null;
        try {
            con = getConnection();
            String student = "create table if not exists student(id integer " + ("sqlite".equals(protocol) ? "" : "auto_increment") + " primary key, name text, age integer, create_date timestamp, update_date timestamp)";
            ps = con.prepareStatement(student);
            ps.executeUpdate();
            ps.close();

            ps = con.prepareStatement("select id from student limit 1");
            rs = ps.executeQuery();
            // No data in database
            if (!rs.next()) {
                rs.close();
                ps.close();
                con.setAutoCommit(false);
                ps = con.prepareStatement("insert into student(name, age, create_date, update_date) values (?,?,?,?)");
                int size = 10_000;
                for (int i = 0; i < size; i++) {
                    ps.setString(1, getRandomString());
                    if (random.nextInt(1000) >= 975) {
                        ps.setNull(2, Types.INTEGER);
                    } else {
                        ps.setInt(2, random.nextInt(15) + 5);
                    }
                    ps.setTimestamp(3, new Timestamp(System.currentTimeMillis()));
                    if (random.nextInt(1000) >= 615) {
                        ps.setNull(4, Types.DATE);
                    } else {
                        ps.setTimestamp(4, new Timestamp(System.currentTimeMillis() - random.nextInt(9999999)));
                    }
                    ps.addBatch();
                }
                ps.executeBatch();
                con.commit();
            } else {
                rs.close();
                ps.close();
            }

            String typesTest = "CREATE TABLE if not exists `types_test` (`id` int(11) NOT NULL,`t_bit` bit(1) DEFAULT NULL,`t_tinyint` tinyint(3) DEFAULT NULL,`t_smallint` smallint(6) DEFAULT NULL,`t_int` int(11) DEFAULT NULL,`t_bigint` bigint(19) DEFAULT NULL,`t_float` float DEFAULT NULL,`t_double` double DEFAULT NULL,`t_varchar` varchar(45) DEFAULT NULL,`t_char` char(10) DEFAULT NULL,`t_date` date DEFAULT NULL,`t_datetime` datetime DEFAULT NULL,`t_timestamp` timestamp NULL DEFAULT NULL)";
            ps = con.prepareStatement(typesTest);
            ps.executeUpdate();
            ps.close();

            ps = con.prepareStatement("select id from types_test limit 1");
            rs = ps.executeQuery();
            // No data in database
            if (!rs.next()) {
                rs.close();
                ps.close();
                con.setAutoCommit(true);
                ps = con.prepareStatement("insert into types_test(id,t_bit,t_tinyint,t_smallint,t_int,t_bigint,t_float,t_double,t_varchar,t_char,t_date,t_datetime,t_timestamp) values(?,?,?,?,?,?,?,?,?,?,?,?,?)");
                ps.setInt(1, 1);
                ps.setByte(2, (byte) 1);
                ps.setByte(3, (byte) 123);
                ps.setShort(4, (short) 1234);
                ps.setInt(5, 12345678);
                ps.setLong(6, 125436456543L);
                ps.setFloat(7, 129.304F);
                ps.setDouble(8, 47484.395D);
                ps.setString(9, "abc");
                ps.setString(10, "are");
                ps.setDate(11, new Date(System.currentTimeMillis()));
                ps.setDate(12, new Date(System.currentTimeMillis()));
                ps.setTimestamp(13, new Timestamp(System.currentTimeMillis()));
                ps.executeUpdate();
                ps = con.prepareStatement("insert into types_test(id) values(?)");
                ps.setInt(1, 2);
                ps.executeUpdate();
                ps.close();
            } else {
                rs.close();
                ps.close();
            }
        } catch (SQLException e) {
            e.printStackTrace();
        } finally {
            if (rs != null) {
                try {
                    rs.close();
                } catch (SQLException e) {
                    e.printStackTrace();
                }
            }
            if (ps != null) {
                try {
                    ps.close();
                } catch (SQLException e) {
                    e.printStackTrace();
                }
            }
            if (con != null) {
                try {
                    con.close();
                } catch (SQLException e) {
                    e.printStackTrace();
                }
            }
        }
    }
}

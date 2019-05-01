/*
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

import org.junit.Before;

import java.io.IOException;
import java.sql.*;
import java.util.Properties;

/**
 * Create by guanquan.wang at 2019-04-28 21:50
 */
public class SQLWorkbookTest extends WorkbookTest {
    private static Properties pro;
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
    }

    protected Connection getConnection() throws SQLException {
        return DriverManager.getConnection(pro.getProperty("dataSource.url"));
    }

    /**
     * Install test data
     */
    @Before public void init() {
        try (Connection con = getConnection()) {
            String student = "create table if not exists student(id integer primary key, name text, age integer)";
            PreparedStatement ps = con.prepareStatement(student);
            ps.executeUpdate();
            ps.close();

            ps = con.prepareStatement("select id from student limit 1");
            ResultSet rs = ps.executeQuery();
            // No data in database
            if (!rs.next()) {
                ps.close();
                con.setAutoCommit(false);
                ps = con.prepareStatement("insert into student(name, age) values (?,?)");
                int size = 10_000;
                for (int i = 0; i < size; i++) {
                    ps.setString(1, getRandomString());
                    ps.setInt(2, random.nextInt(15) + 5);
                    ps.addBatch();
                }
                ps.executeBatch();
                con.commit();
            }
        } catch (SQLException e) {
            e.printStackTrace();
        }
    }
}

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

import cn.ttzero.excel.Print;
import org.junit.Test;

import java.io.IOException;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;

/**
 * Create by guanquan.wang at 2019-04-28 21:50
 */
public class ResultSetSheetTest extends SQLWorkbookTest {
    @Test public void testWrite() {
        try (Connection con = getConnection()) {
            PreparedStatement ps = con.prepareStatement("select id, name, age from student order by age");
            ResultSet rs = ps.executeQuery();
            new Workbook("result set", author)
                .watch(Print::println)
                .setConnection(con)
                .addSheet(rs
                    , new Sheet.Column("学号", int.class)
                    , new Sheet.Column("性名", String.class)
                    , new Sheet.Column("年龄", int.class)
                )
                .writeTo(defaultTestPath);
            ps.close();
        } catch (SQLException|IOException e) {
            e.printStackTrace();
        }
    }
}

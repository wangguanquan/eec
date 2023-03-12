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

import org.junit.Test;
import org.ttzero.excel.Print;
import org.ttzero.excel.entity.style.Fill;
import org.ttzero.excel.entity.style.PatternType;
import org.ttzero.excel.entity.style.Styles;

import java.awt.*;
import java.io.IOException;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Timestamp;

/**
 * @author guanquan.wang at 2019-04-29 15:16
 */
public class ResultSetPagingTest extends SQLWorkbookTest {
    @Test public void testPaging() throws SQLException, IOException {
        try (Connection con = getConnection()) {
            PreparedStatement ps = con.prepareStatement("select id, name, age, create_date, update_date from student");
            ResultSet rs = ps.executeQuery();
            new Workbook("result set paging", author)
                .watch(Print::println)
                .addSheet(rs
                    , new Column("学号", int.class)
                    , new Column("性名", String.class)
                    , new Column("年龄", int.class)
                    , new Column("创建时间", Timestamp.class)
                    , new Column("更新", Timestamp.class)
                )
            .setWorkbookWriter(new ReLimitXMLWorkbookWriter())
            .writeTo(defaultTestPath);
            ps.close();
        }
    }


    @Test public void testStyleDesignPaging() throws SQLException, IOException {
        try (Connection con = getConnection()) {
            PreparedStatement ps = con.prepareStatement("select id, name, age, create_date, update_date from student");
            ResultSet rs = ps.executeQuery();
            new Workbook("test global style design for ResultSet Paging", author)
                .watch(Print::println)
                .addSheet(new ResultSetSheet().setRs(rs).setStyleProcessor((rst, style, sst)->{
                    try {
                        if (rst.getInt("age") > 14) {
                            style = Styles.clearFill(style) | sst.addFill(new Fill(PatternType.solid, Color.yellow));
                        }
                    } catch (SQLException ex) {
                        // Ignore
                    }
                    return style;
                }))
                .setWorkbookWriter(new ReLimitXMLWorkbookWriter())
                .writeTo(defaultTestPath);
            ps.close();
        }
    }
}

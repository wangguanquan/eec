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

import java.awt.Color;
import java.io.IOException;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;

/**
 * @author guanquan.wang at 2019-05-01 19:34
 */
public class MultiWorksheetTest extends SQLWorkbookTest {

    @Test
    public void testMultiWorksheet() throws IOException {
        new Workbook("test multi worksheet", author)
                .watch(Print::println)
                .setAutoSize(true)
                // The first worksheet
                .addSheet("E", ListMapSheetTest.createTestData())
                // The other worksheet
                .addSheet("All type", ListMapSheetTest.createAllTypeData())
                .writeTo(defaultTestPath);
    }

    @Test
    public void testMultiDataSource() throws SQLException, IOException {
        try (
            Connection con = getConnection();
            PreparedStatement ps = con.prepareStatement("select id, name, age from student order by age limit 10");
            ResultSet rs = ps.executeQuery()
        ) {

            new Workbook("test multi dataSource worksheet", author)
                .watch(Print::println)
                .setAutoSize(true)
                .setConnection(con)
                // List<Map>
                .addSheet("ListMap", ListMapSheetTest.createAllTypeData())
                // List<Object>
                .addSheet("ListObject", ListObjectSheetTest.Item.randomTestData())
                // Statement
                .addSheet("Statement", "select id, name, age from student"
                    , new Sheet.Column("学号", int.class)
                    , new Sheet.Column("姓名", String.class)
                    , new Sheet.Column("年龄", int.class, n -> n > 14 ? "高龄" : n)
                        .setStyleProcessor((o, style, sst) -> {
                            int n = (int) o;
                            if (n > 14) {
                                style = Styles.clearFill(style)
                                    | sst.addFill(new Fill(PatternType.solid, Color.orange));
                            }
                            return style;
                        })
                )
                // Empty
                .addSheet(new EmptySheet("Empty"))
                // ResultSet
                .addSheet("ResultSet", rs
                    , new Sheet.Column("学号", int.class)
                    , new Sheet.Column("姓名", String.class)
                    , new Sheet.Column("年龄", int.class)
                )
                // Customize
                .addSheet(new CustomizeDataSourceSheet("Customize"))
                .writeTo(defaultTestPath);
        }
    }
}

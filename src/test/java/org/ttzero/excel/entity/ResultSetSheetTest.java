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
 * @author guanquan.wang at 2019-04-28 21:50
 */
public class ResultSetSheetTest extends SQLWorkbookTest {
    @Test public void testWrite() throws SQLException, IOException {
        try (
            Connection con = getConnection();
            PreparedStatement ps = con.prepareStatement("select id, name, age, create_date, update_date from student limit 10");
            ResultSet rs = ps.executeQuery()
        ) {
            new Workbook("result set", author)
                .watch(Print::println)
                .addSheet(rs
                    , new Column("学号", int.class)
                    , new Column("性名", String.class)
                    , new Column("年龄", Integer.class)
                    , new Column("创建时间", Timestamp.class)
                    , new Column("更新", Timestamp.class)
                )
                .writeTo(defaultTestPath);
        }
    }

    @Test public void testStyleDesign4RS() throws IOException, SQLException {
        try (
            Connection con = getConnection();
            PreparedStatement ps = con.prepareStatement("select id, name, age, create_date, update_date from student limit 10");
            ResultSet rs = ps.executeQuery()
        ) {
            new Workbook("test global style design for ResultSet", author)
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
                .writeTo(defaultTestPath);
        }
    }

    @Test public void testConstructor1() throws IOException {
        try {
            new Workbook("test ResultSet sheet Constructor1", author)
                    .watch(Print::println)
                    .addSheet(new ResultSetSheet())
                    .writeTo(defaultTestPath);
        } catch (ExcelWriteException e) {
            assert true;
        }
    }

    @Test public void testConstructor2() throws SQLException, IOException {
        try (
            Connection con = getConnection();
            PreparedStatement ps = con.prepareStatement("select id, name, age, create_date, update_date from student limit 10");
            ResultSet rs = ps.executeQuery()
        ) {
            new Workbook("test ResultSet sheet Constructor2", author)
                .watch(Print::println)
                .addSheet(new ResultSetSheet().setRs(rs))
                .writeTo(defaultTestPath);
        }
    }

    @Test public void testConstructor3() throws SQLException, IOException {
        try (
            Connection con = getConnection();
            PreparedStatement ps = con.prepareStatement("select id, name, age, create_date, update_date from student limit 10");
            ResultSet rs = ps.executeQuery()
        ) {
            new Workbook("test ResultSet sheet Constructor3", author)
                .watch(Print::println)
                .addSheet(new ResultSetSheet("Student").setRs(rs))
                .writeTo(defaultTestPath);
        }
    }

    @Test public void testConstructor4() throws SQLException, IOException {
        try (
            Connection con = getConnection();
            PreparedStatement ps = con.prepareStatement("select id, name, age, create_date, update_date from student limit 10");
            ResultSet rs = ps.executeQuery()
        ) {
            new Workbook("test ResultSet sheet Constructor4", author)
                .watch(Print::println)
                .addSheet(new ResultSetSheet("Student"
                    , new Column("ID", int.class)
                    , new Column("NAME", String.class)
                    , new Column("AGE", int.class)
                    , new Column("CREATE_DATE", Timestamp.class)
                    , new Column("UPDATE_DATE", Timestamp.class)
                ).setRs(rs))
                .writeTo(defaultTestPath);
        }
    }

    @Test public void testConstructor5() throws SQLException, IOException {
        try (
            Connection con = getConnection();
            PreparedStatement ps = con.prepareStatement("select id, name, age, create_date, update_date from student limit 10");
            ResultSet rs = ps.executeQuery()
        ) {
            new Workbook("test ResultSet sheet Constructor5", author)
                .watch(Print::println)
                .addSheet(new ResultSetSheet("Student", WaterMark.of(author)
                    , new Column("ID", int.class)
                    , new Column("NAME", String.class)
                    , new Column("AGE", int.class)
                    , new Column("CREATE_DATE", Timestamp.class)
                    , new Column("UPDATE_DATE", Timestamp.class)
                ).setRs(rs))
                .writeTo(defaultTestPath);
        }
    }

    @Test public void testConstructor6() throws SQLException, IOException {
        try (
            Connection con = getConnection();
            PreparedStatement ps = con.prepareStatement("select id, name, age, create_date, update_date from student limit 10");
            ResultSet rs = ps.executeQuery()
        ) {
            new Workbook("test ResultSet sheet Constructor6", author)
                .watch(Print::println)
                .addSheet(new ResultSetSheet(rs))
                .writeTo(defaultTestPath);
        }
    }

    @Test public void testConstructor7() throws SQLException, IOException {
        try (
            Connection con = getConnection();
            PreparedStatement ps = con.prepareStatement("select id, name, age, create_date, update_date from student limit 10");
            ResultSet rs = ps.executeQuery()
        ) {
            new Workbook("test ResultSet sheet Constructor7", author)
                .watch(Print::println)
                .addSheet(new ResultSetSheet("Student", rs))
                .writeTo(defaultTestPath);
        }
    }

    @Test public void testConstructor8() throws SQLException, IOException {
        try (
            Connection con = getConnection();
            PreparedStatement ps = con.prepareStatement("select id, name, age, create_date, update_date from student limit 10");
            ResultSet rs = ps.executeQuery()
        ) {
            new Workbook("test ResultSet sheet Constructor8", author)
                .watch(Print::println)
                .addSheet(new ResultSetSheet(rs
                    , new Column("ID", int.class)
                    , new Column("NAME", String.class)
                    , new Column("AGE", int.class)
                    , new Column("CREATE_DATE", Timestamp.class)
                    , new Column("UPDATE_DATE", Timestamp.class)
                ))
                .writeTo(defaultTestPath);
        }
    }

    @Test public void testConstructor9() throws SQLException, IOException {
        try (
            Connection con = getConnection();
            PreparedStatement ps = con.prepareStatement("select id, name, age, create_date, update_date from student limit 10");
            ResultSet rs = ps.executeQuery()
        ) {
            new Workbook("test ResultSet sheet Constructor9", author)
                .watch(Print::println)
                .addSheet(new ResultSetSheet("Student", rs
                    , new Column("ID", int.class)
                    , new Column("NAME", String.class)
                    , new Column("AGE", int.class)
                    , new Column("CREATE_DATE", Timestamp.class)
                    , new Column("UPDATE_DATE", Timestamp.class)
                ))
                .writeTo(defaultTestPath);
        }
    }

    @Test public void testConstructor10() throws SQLException, IOException {
        try (
            Connection con = getConnection();
            PreparedStatement ps = con.prepareStatement("select id, name, age, create_date, update_date from student limit 10");
            ResultSet rs = ps.executeQuery()
        ) {
            new Workbook("test ResultSet sheet Constructor10", author)
                .watch(Print::println)
                .addSheet(new ResultSetSheet(rs, WaterMark.of(author)
                    , new Column("ID", int.class)
                    , new Column("NAME", String.class)
                    , new Column("AGE", int.class)
                    , new Column("CREATE_DATE", Timestamp.class)
                    , new Column("UPDATE_DATE", Timestamp.class)
                ))
                .writeTo(defaultTestPath);
        }
    }

    @Test public void testConstructor11() throws SQLException, IOException {
        try (
            Connection con = getConnection();
            PreparedStatement ps = con.prepareStatement("select id, name, age, create_date, update_date from student limit 10");
            ResultSet rs = ps.executeQuery()
        ) {
            new Workbook("test ResultSet sheet Constructor11", author)
                .watch(Print::println)
                .addSheet(new ResultSetSheet("Student", rs, WaterMark.of(author)
                    , new Column("ID", int.class)
                    , new Column("NAME", String.class)
                    , new Column("AGE", int.class)
                    , new Column("AGE", int.class)
                    , new Column("UPDATE_DATE", Timestamp.class)
                ))
                .writeTo(defaultTestPath);
        }
    }

    @Test public void testDiffTypeFromMetadata() throws SQLException, IOException {
        try (
            Connection con = getConnection();
            PreparedStatement ps = con.prepareStatement("select id, name, age, create_date, update_date from student limit 10");
            ResultSet rs = ps.executeQuery()
        ) {
            new Workbook("test ResultSet different type from metadata", author)
                .watch(Print::println)
                .addSheet(new ResultSetSheet("Student", rs, WaterMark.of(author)
                    , new Column("ID", String.class)  // Integer in database
                    , new Column("NAME", String.class)
                    , new Column("AGE", String.class) // // Integer in database
                    , new Column("AGE", int.class)
                    , new Column("UPDATE_DATE", Timestamp.class)
                ))
                .writeTo(defaultTestPath);
        }
    }


}

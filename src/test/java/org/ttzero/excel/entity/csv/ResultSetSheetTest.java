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

package org.ttzero.excel.entity.csv;

import org.junit.Test;
import org.ttzero.excel.Print;
import org.ttzero.excel.entity.Column;
import org.ttzero.excel.entity.ExcelWriteException;
import org.ttzero.excel.entity.ResultSetSheet;
import org.ttzero.excel.entity.SQLWorkbookTest;
import org.ttzero.excel.entity.Sheet;
import org.ttzero.excel.entity.Workbook;

import java.io.IOException;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;

/**
 * @author guanquan.wang at 2019-04-28 21:50
 */
public class ResultSetSheetTest extends SQLWorkbookTest {
    @Test public void testWrite() throws SQLException, IOException {
        try (
            Connection con = getConnection();
            PreparedStatement ps = con.prepareStatement("select id, name, age from student limit 10");
            ResultSet rs = ps.executeQuery()
        ) {
            new Workbook("result set")
                .watch(Print::println)
                .setConnection(con)
                .addSheet(rs
                    , new Column("学号", int.class)
                    , new Column("性名", String.class)
                    , new Column("年龄", int.class)
                )
                .saveAsCSV()
                .writeTo(getOutputTestPath());
        }
    }

    @Test public void testConstructor1() throws IOException {
        try {
            new Workbook("test ResultSet sheet Constructor1")
                .watch(Print::println)
                .addSheet(new ResultSetSheet())
                .saveAsCSV()
                .writeTo(getOutputTestPath());
        } catch (ExcelWriteException e) {
            assert true;
        }
    }

    @Test public void testConstructor2() throws SQLException, IOException {
        try (
            Connection con = getConnection();
            PreparedStatement ps = con.prepareStatement("select id, name, age from student limit 10");
            ResultSet rs = ps.executeQuery()
        ) {
            new Workbook("test ResultSet sheet Constructor2")
                .watch(Print::println)
                .addSheet(new ResultSetSheet().setRs(rs))
                .saveAsCSV()
                .writeTo(getOutputTestPath());
        }
    }

    @Test public void testConstructor3() throws SQLException, IOException {
        try (
            Connection con = getConnection();
            PreparedStatement ps = con.prepareStatement("select id, name, age from student limit 10");
            ResultSet rs = ps.executeQuery()
        ) {
            new Workbook("test ResultSet sheet Constructor3")
                .watch(Print::println)
                .addSheet(new ResultSetSheet("Student").setRs(rs))
                .saveAsCSV()
                .writeTo(getOutputTestPath());
        }
    }

    @Test public void testConstructor4() throws SQLException, IOException {
        try (
            Connection con = getConnection();
            PreparedStatement ps = con.prepareStatement("select id, name, age from student limit 10");
            ResultSet rs = ps.executeQuery()
        ) {
            new Workbook("test ResultSet sheet Constructor4")
                .watch(Print::println)
                .addSheet(new ResultSetSheet("Student"
                    , new Column("ID", int.class)
                    , new Column("NAME", String.class)
                    , new Column("AGE", int.class))
                    .setRs(rs))
                .saveAsCSV()
                .writeTo(getOutputTestPath());
        }
    }

    @Test public void testConstructor5() throws SQLException, IOException {
        try (
            Connection con = getConnection();
            PreparedStatement ps = con.prepareStatement("select id, name, age from student limit 10");
            ResultSet rs = ps.executeQuery()
        ) {
            new Workbook("test ResultSet sheet Constructor5")
                .watch(Print::println)
                .addSheet(new ResultSetSheet("Student"
                    , new Column("ID", int.class)
                    , new Column("NAME", String.class)
                    , new Column("AGE", int.class))
                    .setRs(rs))
                .saveAsCSV()
                .writeTo(getOutputTestPath());
        }
    }

    @Test public void testConstructor6() throws SQLException, IOException {
        try (
            Connection con = getConnection();
            PreparedStatement ps = con.prepareStatement("select id, name, age from student limit 10");
            ResultSet rs = ps.executeQuery()
        ) {
            new Workbook("test ResultSet sheet Constructor6")
                .watch(Print::println)
                .addSheet(new ResultSetSheet(rs))
                .saveAsCSV()
                .writeTo(getOutputTestPath());
        }
    }

    @Test public void testConstructor7() throws SQLException, IOException {
        try (
            Connection con = getConnection();
            PreparedStatement ps = con.prepareStatement("select id, name, age from student limit 10");
            ResultSet rs = ps.executeQuery()
        ) {
            new Workbook("test ResultSet sheet Constructor7")
                .watch(Print::println)
                .addSheet(new ResultSetSheet("Student", rs))
                .saveAsCSV()
                .writeTo(getOutputTestPath());
        }
    }

    @Test public void testConstructor8() throws SQLException, IOException {
        try (
            Connection con = getConnection();
            PreparedStatement ps = con.prepareStatement("select id, name, age from student limit 10");
            ResultSet rs = ps.executeQuery()
        ) {
            new Workbook("test ResultSet sheet Constructor8")
                .watch(Print::println)
                .addSheet(new ResultSetSheet(rs
                    , new Column("ID", int.class)
                    , new Column("NAME", String.class)
                    , new Column("AGE", int.class)
                ))
                .saveAsCSV()
                .writeTo(getOutputTestPath());
        }
    }

    @Test public void testConstructor9() throws SQLException, IOException {
        try (
            Connection con = getConnection();
            PreparedStatement ps = con.prepareStatement("select id, name, age from student limit 10");
            ResultSet rs = ps.executeQuery()
        ) {
            new Workbook("test ResultSet sheet Constructor9")
                .watch(Print::println)
                .addSheet(new ResultSetSheet("Student", rs
                    , new Column("ID", int.class)
                    , new Column("NAME", String.class)
                    , new Column("AGE", int.class)
                ))
                .saveAsCSV()
                .writeTo(getOutputTestPath());
        }
    }

    @Test public void testConstructor10() throws SQLException, IOException {
        try (
            Connection con = getConnection();
            PreparedStatement ps = con.prepareStatement("select id, name, age from student limit 10");
            ResultSet rs = ps.executeQuery()
        ) {
            new Workbook("test ResultSet sheet Constructor10")
                .watch(Print::println)
                .addSheet(new ResultSetSheet(rs
                    , new Column("ID", int.class)
                    , new Column("NAME", String.class)
                    , new Column("AGE", int.class)
                ))
                .saveAsCSV()
                .writeTo(getOutputTestPath());
        }
    }

    @Test public void testConstructor11() throws SQLException, IOException {
        try (
            Connection con = getConnection();
            PreparedStatement ps = con.prepareStatement("select id, name, age from student limit 10");
            ResultSet rs = ps.executeQuery()
        ) {
            new Workbook("test ResultSet sheet Constructor11")
                .watch(Print::println)
                .addSheet(new ResultSetSheet("Student", rs
                    , new Column("ID", int.class)
                    , new Column("NAME", String.class)
                    , new Column("AGE", int.class)
                ))
                .saveAsCSV()
                .writeTo(getOutputTestPath());
        }
    }

    @Test public void testDiffTypeFromMetadata() throws SQLException, IOException {
        try (
            Connection con = getConnection();
            PreparedStatement ps = con.prepareStatement("select id, name, age from student limit 10");
            ResultSet rs = ps.executeQuery()
        ) {
            new Workbook("test ResultSet different type from metadata")
                .watch(Print::println)
                .addSheet(new ResultSetSheet("Student", rs
                    , new Column("ID", String.class)  // Integer in database
                    , new Column("NAME", String.class)
                    , new Column("AGE", String.class) // // Integer in database
                ))
                .saveAsCSV()
                .writeTo(getOutputTestPath());
        }
    }


}

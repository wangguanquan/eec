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
import org.ttzero.excel.entity.SQLWorkbookTest;
import org.ttzero.excel.entity.Sheet;
import org.ttzero.excel.entity.StatementSheet;
import org.ttzero.excel.entity.Workbook;
import org.ttzero.excel.entity.style.Fill;
import org.ttzero.excel.entity.style.PatternType;
import org.ttzero.excel.entity.style.Styles;

import java.awt.Color;
import java.io.IOException;
import java.sql.Connection;
import java.sql.SQLException;

/**
 * @author guanquan.wang at 2019-04-28 22:47
 */
public class StatementSheetTest extends SQLWorkbookTest {
    @Test public void testWrite() throws SQLException, IOException {
        try (Connection con = getConnection()) {
            new Workbook("statement")
                .watch(Print::println)
                .setConnection(con)
                .addSheet("select id, name, age from student order by age"
                    , new Column("学号", int.class)
                    , new Column("性名", String.class)
                    , new Column("年龄", int.class)
                )
                .saveAsCSV()
                .writeTo(getOutputTestPath());
        }
    }

    @Test public void testStyleProcessor() throws SQLException, IOException {
        try (Connection con = getConnection()) {
            new Workbook("statement style processor")
                .watch(Print::println)
                .setConnection(con)
                .addSheet("select id, name, age from student"
                    , new Column("学号", int.class)
                    , new Column("性名", String.class)
                    , new Column("年龄", int.class)
                        .setStyleProcessor((o, style, sst) -> {
                            int n = (int) o;
                            if (n < 10) {
                                style = Styles.clearFill(style)
                                    | sst.addFill(new Fill(PatternType.solid, Color.orange));
                            }
                            return style;
                        })
                )
                .saveAsCSV()
                .writeTo(getOutputTestPath());
        }
    }

    @Test public void testIntConversion() throws SQLException, IOException {
        try (Connection con = getConnection()) {
            new Workbook("test int conversion statement")
                .setConnection(con)
                .watch(Print::println)
                .addSheet("select id, name, age from student"
                    , new Column("学号", int.class)
                    , new Column("姓名", String.class)
                    , new Column("年龄", int.class, n -> n > 14 ? "高龄" : n)
                        .setStyleProcessor((o, style, sst) -> {
                            int n = (int) o;
                            if (n > 14) {
                                style = Styles.clearFill(style)
                                    | sst.addFill(new Fill(PatternType.solid, Color.orange));
                            }
                            return style;
                        })
                )
                .saveAsCSV()
                .writeTo(getOutputTestPath());
        }
    }

    @Test public void testConstructor1() throws SQLException, IOException {
        try (Connection con = getConnection()) {
            new Workbook("test statement sheet Constructor1")
                .watch(Print::println)
                .addSheet(new StatementSheet(con, "select id, name, age from student limit 10"))
                .saveAsCSV()
                .writeTo(getOutputTestPath());
        }
    }

    @Test public void testConstructor2() throws SQLException, IOException {
        try (Connection con = getConnection()) {
            new Workbook("test statement sheet Constructor2")
                .watch(Print::println)
                .addSheet(new StatementSheet("Student", con, "select id, name, age from student limit 10"))
                .saveAsCSV()
                .writeTo(getOutputTestPath());
        }
    }

    @Test public void testConstructor3() throws SQLException, IOException {
        try (Connection con = getConnection()) {
            new Workbook("test statement sheet Constructor3")
                .watch(Print::println)
                .addSheet(new StatementSheet(con, "select id, name, age from student where id between ? and ?", ps -> {
                    ps.setInt(1, 10);
                    ps.setInt(2, 20);
                }))
                .saveAsCSV()
                .writeTo(getOutputTestPath());
        }
    }

    @Test public void testConstructor4() throws SQLException, IOException {
        try (Connection con = getConnection()) {
            new Workbook("test statement sheet Constructor4")
                .watch(Print::println)
                .addSheet(new StatementSheet("Student", con, "select id, name, age from student where id between ? and ?", ps -> {
                    ps.setInt(1, 10);
                    ps.setInt(2, 20);
                }))
                .saveAsCSV()
                .writeTo(getOutputTestPath());
        }
    }

    @Test public void testConstructor5() throws SQLException, IOException {
        try (Connection con = getConnection()) {
            new Workbook("test statement sheet Constructor5")
                .watch(Print::println)
                .addSheet(new StatementSheet(con, "select id, name, age from student limit 10"
                    , new Column("ID", int.class)
                    , new Column("NAME", String.class)
                    , new Column("AGE", int.class)
                ))
                .saveAsCSV()
                .writeTo(getOutputTestPath());
        }
    }

    @Test public void testConstructor6() throws SQLException, IOException {
        try (Connection con = getConnection()) {
            new Workbook("test statement sheet Constructor6")
                .watch(Print::println)
                .addSheet(new StatementSheet("Student", con, "select id, name, age from student limit 10"
                    , new Column("ID", int.class)
                    , new Column("NAME", String.class)
                    , new Column("AGE", int.class)
                ))
                .saveAsCSV()
                .writeTo(getOutputTestPath());
        }
    }

    @Test public void testConstructor7() throws SQLException, IOException {
        try (Connection con = getConnection()) {
            new Workbook("test statement sheet Constructor7")
                .watch(Print::println)
                .addSheet(new StatementSheet(con, "select id, name, age from student where id between ? and ?"
                    , ps -> {
                        ps.setInt(1, 10);
                        ps.setInt(2, 20);
                    }
                    , new Column("ID", int.class)
                    , new Column("NAME", String.class)
                    , new Column("AGE", int.class)))
                .saveAsCSV()
                .writeTo(getOutputTestPath());
        }
    }

    @Test public void testConstructor8() throws SQLException, IOException {
        try (Connection con = getConnection()) {
            new Workbook("test statement sheet Constructor8")
                .watch(Print::println)
                .addSheet(new StatementSheet("Student", con, "select id, name, age from student where id between ? and ?"
                    , ps -> {
                        ps.setInt(1, 10);
                        ps.setInt(2, 20);
                    }
                    , new Column("ID", int.class)
                    , new Column("NAME", String.class)
                    , new Column("AGE", int.class)))
                .saveAsCSV()
                .writeTo(getOutputTestPath());
        }
    }

    @Test public void testConstructor9() throws IOException {
        try {
            new Workbook("test statement sheet Constructor9")
                .watch(Print::println)
                .addSheet(new StatementSheet())
                .saveAsCSV()
                .writeTo(getOutputTestPath());
        } catch (ExcelWriteException e) {
            assert true;
        }
    }

    @Test public void testConstructor10() throws SQLException, IOException {
        try (Connection con = getConnection()) {
            new Workbook("test statement sheet Constructor10")
                .watch(Print::println)
                .addSheet(new StatementSheet()
                    .setPs(con.prepareStatement("select id, name, age from student limit 10")))
                .saveAsCSV()
                .writeTo(getOutputTestPath());
        }
    }

    @Test public void testConstructor11() throws SQLException, IOException {
        try (Connection con = getConnection()) {
            new Workbook("test statement sheet Constructor11")
                .watch(Print::println)
                .addSheet(new StatementSheet("Student")
                    .setPs(con.prepareStatement("select id, name, age from student limit 10")))
                .saveAsCSV()
                .writeTo(getOutputTestPath());
        }
    }

    @Test public void testConstructor12() throws SQLException, IOException {
        try (Connection con = getConnection()) {
            new Workbook("test statement sheet Constructor12")
                .watch(Print::println)
                .addSheet(new StatementSheet("Student")
                    .setPs(con.prepareStatement("select id, name, age from student limit 10")))
                .saveAsCSV()
                .writeTo(getOutputTestPath());
        }
    }

    @Test public void testConstructor13() throws SQLException, IOException {
        try (Connection con = getConnection()) {
            new Workbook("test statement sheet Constructor13")
                .watch(Print::println)
                .addSheet(new StatementSheet("Student"
                    , new Column("ID", int.class)
                    , new Column("NAME", String.class)
                    , new Column("AGE", int.class))
                    .setPs(con.prepareStatement("select id, name, age from student limit 10")))
                .saveAsCSV()
                .writeTo(getOutputTestPath());
        }
    }


    @Test public void testCancelOddStyle() throws SQLException, IOException {
        try (Connection con = getConnection()) {
            new Workbook("test statement sheet cancel odd")
                .watch(Print::println)
                .addSheet(new StatementSheet(con, "select id, name, age from student limit 10")
                    .cancelOddStyle()
                )
                .saveAsCSV()
                .writeTo(getOutputTestPath());
        }
    }

    @Test public void testDiffTypeFromMetadata() throws SQLException, IOException {
        try (Connection con = getConnection()) {
            new Workbook("test Statement different type from metadata")
                .watch(Print::println)
                .addSheet(new StatementSheet(con, "select id, name, age from student limit 10"
                    , new Column("ID", String.class)  // Integer in database
                    , new Column("NAME", String.class)
                    , new Column("AGE", String.class) // Integer in database
                ))
                .saveAsCSV()
                .writeTo(getOutputTestPath());
        }
    }

    @Test public void testFixWidth() throws SQLException, IOException {
        try (Connection con = getConnection()) {
            new Workbook("test statement fix width")
                .watch(Print::println)
                .addSheet(new StatementSheet(con, "select id, name, age from student limit 10").fixSize(10))
                .saveAsCSV()
                .writeTo(getOutputTestPath());
        }
    }
}

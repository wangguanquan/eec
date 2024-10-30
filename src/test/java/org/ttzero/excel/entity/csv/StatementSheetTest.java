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
import org.ttzero.excel.entity.Column;
import org.ttzero.excel.entity.ExcelWriteException;
import org.ttzero.excel.entity.SQLWorkbookTest;
import org.ttzero.excel.entity.StatementSheet;
import org.ttzero.excel.entity.Workbook;
import org.ttzero.excel.entity.style.Fill;
import org.ttzero.excel.entity.style.PatternType;
import org.ttzero.excel.entity.style.Styles;

import java.awt.Color;
import java.io.IOException;
import java.sql.Connection;
import java.sql.SQLException;

import static org.junit.Assert.assertThrows;

/**
 * @author guanquan.wang at 2019-04-28 22:47
 */
public class StatementSheetTest extends SQLWorkbookTest {
    @Test public void testWrite() throws SQLException, IOException {
        try (Connection con = getConnection()) {
            new Workbook()
                .addSheet(new StatementSheet(con, "select id, name, age from student order by age"
                    , new Column("学号", int.class)
                    , new Column("性名", String.class)
                    , new Column("年龄", int.class)
                ))
                .saveAsCSV()
                .writeTo(getOutputTestPath().resolve("statement.csv"));
        }
    }

    @Test public void testStyleProcessor() throws SQLException, IOException {
        try (Connection con = getConnection()) {
            new Workbook()
                .addSheet(new StatementSheet(con, "select id, name, age from student"
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
                ))
                .saveAsCSV()
                .writeTo(getOutputTestPath().resolve("statement style processor.csv"));
        }
    }

    @Test public void testIntConversion() throws SQLException, IOException {
        try (Connection con = getConnection()) {
            new Workbook()
                .addSheet(new StatementSheet(con, "select id, name, age from student"
                    , new Column("学号", int.class)
                    , new Column("姓名", String.class)
                    , new Column("年龄", int.class, n -> (int) n > 14 ? "高龄" : n)
                        .setStyleProcessor((o, style, sst) -> {
                            int n = (int) o;
                            if (n > 14) {
                                style = Styles.clearFill(style)
                                    | sst.addFill(new Fill(PatternType.solid, Color.orange));
                            }
                            return style;
                        })
                ))
                .saveAsCSV()
                .writeTo(getOutputTestPath().resolve("test int conversion statement.csv"));
        }
    }

    @Test public void testConstructor1() throws SQLException, IOException {
        try (Connection con = getConnection()) {
            new Workbook()
                .addSheet(new StatementSheet(con, "select id, name, age from student limit 10"))
                .saveAsCSV()
                .writeTo(getOutputTestPath().resolve("test statement sheet Constructor1.csv"));
        }
    }

    @Test public void testConstructor2() throws SQLException, IOException {
        try (Connection con = getConnection()) {
            new Workbook()
                .addSheet(new StatementSheet("Student", con, "select id, name, age from student limit 10"))
                .saveAsCSV()
                .writeTo(getOutputTestPath().resolve("test statement sheet Constructor2.csv"));
        }
    }

    @Test public void testConstructor3() throws SQLException, IOException {
        try (Connection con = getConnection()) {
            new Workbook()
                .addSheet(new StatementSheet(con, "select id, name, age from student where id between ? and ?", ps -> {
                    ps.setInt(1, 10);
                    ps.setInt(2, 20);
                }))
                .saveAsCSV()
                .writeTo(getOutputTestPath().resolve("test statement sheet Constructor3.csv"));
        }
    }

    @Test public void testConstructor4() throws SQLException, IOException {
        try (Connection con = getConnection()) {
            new Workbook()
                .addSheet(new StatementSheet("Student", con, "select id, name, age from student where id between ? and ?", ps -> {
                    ps.setInt(1, 10);
                    ps.setInt(2, 20);
                }))
                .saveAsCSV()
                .writeTo(getOutputTestPath().resolve("test statement sheet Constructor4.csv"));
        }
    }

    @Test public void testConstructor5() throws SQLException, IOException {
        try (Connection con = getConnection()) {
            new Workbook()
                .addSheet(new StatementSheet(con, "select id, name, age from student limit 10"
                    , new Column("ID", int.class)
                    , new Column("NAME", String.class)
                    , new Column("AGE", int.class)
                ))
                .saveAsCSV()
                .writeTo(getOutputTestPath().resolve("test statement sheet Constructor5.csv"));
        }
    }

    @Test public void testConstructor6() throws SQLException, IOException {
        try (Connection con = getConnection()) {
            new Workbook()
                .addSheet(new StatementSheet("Student", con, "select id, name, age from student limit 10"
                    , new Column("ID", int.class)
                    , new Column("NAME", String.class)
                    , new Column("AGE", int.class)
                ))
                .saveAsCSV()
                .writeTo(getOutputTestPath().resolve("test statement sheet Constructor6.csv"));
        }
    }

    @Test public void testConstructor7() throws SQLException, IOException {
        try (Connection con = getConnection()) {
            new Workbook()
                .addSheet(new StatementSheet(con, "select id, name, age from student where id between ? and ?"
                    , ps -> {
                        ps.setInt(1, 10);
                        ps.setInt(2, 20);
                    }
                    , new Column("ID", int.class)
                    , new Column("NAME", String.class)
                    , new Column("AGE", int.class)))
                .saveAsCSV()
                .writeTo(getOutputTestPath().resolve("test statement sheet Constructor7.csv"));
        }
    }

    @Test public void testConstructor8() throws SQLException, IOException {
        try (Connection con = getConnection()) {
            new Workbook()
                .addSheet(new StatementSheet("Student", con, "select id, name, age from student where id between ? and ?"
                    , ps -> {
                        ps.setInt(1, 10);
                        ps.setInt(2, 20);
                    }
                    , new Column("ID", int.class)
                    , new Column("NAME", String.class)
                    , new Column("AGE", int.class)))
                .saveAsCSV()
                .writeTo(getOutputTestPath().resolve("test statement sheet Constructor8.csv"));
        }
    }

    @Test public void testConstructor9() {
        assertThrows(ExcelWriteException.class, () -> new Workbook()
            .addSheet(new StatementSheet())
            .saveAsCSV()
            .writeTo(getOutputTestPath().resolve("test statement sheet Constructor9.csv")));
    }

    @Test public void testConstructor10() throws SQLException, IOException {
        try (Connection con = getConnection()) {
            new Workbook()
                .addSheet(new StatementSheet()
                    .setStatement(con.prepareStatement("select id, name, age from student limit 10")))
                .saveAsCSV()
                .writeTo(getOutputTestPath().resolve("test statement sheet Constructor10.csv"));
        }
    }

    @Test public void testConstructor11() throws SQLException, IOException {
        try (Connection con = getConnection()) {
            new Workbook()
                .addSheet(new StatementSheet("Student")
                    .setStatement(con.prepareStatement("select id, name, age from student limit 10")))
                .saveAsCSV()
                .writeTo(getOutputTestPath().resolve("test statement sheet Constructor11.csv"));
        }
    }

    @Test public void testConstructor12() throws SQLException, IOException {
        try (Connection con = getConnection()) {
            new Workbook()
                .addSheet(new StatementSheet("Student")
                    .setStatement(con.prepareStatement("select id, name, age from student limit 10")))
                .saveAsCSV()
                .writeTo(getOutputTestPath().resolve("test statement sheet Constructor12.csv"));
        }
    }

    @Test public void testConstructor13() throws SQLException, IOException {
        try (Connection con = getConnection()) {
            new Workbook()
                .addSheet(new StatementSheet("Student"
                    , new Column("ID", int.class)
                    , new Column("NAME", String.class)
                    , new Column("AGE", int.class))
                    .setStatement(con.prepareStatement("select id, name, age from student limit 10")))
                .saveAsCSV()
                .writeTo(getOutputTestPath().resolve("test statement sheet Constructor13.csv"));
        }
    }


    @Test public void testCancelOddStyle() throws SQLException, IOException {
        try (Connection con = getConnection()) {
            new Workbook()
                .addSheet(new StatementSheet(con, "select id, name, age from student limit 10")
                    .cancelZebraLine()
                )
                .saveAsCSV()
                .writeTo(getOutputTestPath().resolve("test statement sheet cancel odd.csv"));
        }
    }

    @Test public void testDiffTypeFromMetadata() throws SQLException, IOException {
        try (Connection con = getConnection()) {
            new Workbook()
                .addSheet(new StatementSheet(con, "select id, name, age from student limit 10"
                    , new Column("ID", String.class)  // Integer in database
                    , new Column("NAME", String.class)
                    , new Column("AGE", String.class) // Integer in database
                ))
                .saveAsCSV()
                .writeTo(getOutputTestPath().resolve("test Statement different type from metadata.csv"));
        }
    }

    @Test public void testFixWidth() throws SQLException, IOException {
        try (Connection con = getConnection()) {
            new Workbook()
                .addSheet(new StatementSheet(con, "select id, name, age from student limit 10").fixedSize(10))
                .saveAsCSV()
                .writeTo(getOutputTestPath().resolve("test statement fix width.csv"));
        }
    }
}

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
import java.sql.SQLException;
import java.sql.Timestamp;

/**
 * @author guanquan.wang at 2019-04-28 22:47
 */
public class StatementSheetTest extends SQLWorkbookTest {
    @Test public void testWrite() throws SQLException, IOException {
        testWrite(false);
    }

    @Test public void testStyleProcessor() throws SQLException, IOException {
        testStyleProcessor(false);
    }

    @Test public void testIntConversion() throws SQLException, IOException {
        testIntConversion(false);
    }

    // ---- AUTO SIZE

    @Test public void testWriteAutoSize() throws SQLException, IOException {
        testWrite(true);
    }

    @Test public void testStyleProcessorAutoSize() throws SQLException, IOException {
        testStyleProcessor(true);
    }

    @Test public void testIntConversionAutoSize() throws SQLException, IOException {
        testIntConversion(true);
    }

    private void testWrite(boolean autoSize) throws SQLException, IOException {
        try (Connection con = getConnection()) {
            new Workbook("statement", author)
                .setAutoSize(autoSize)
                .setConnection(con)
                .addSheet("select id, name, age, create_date, update_date from student order by age"
                    , new Column("学号", int.class)
                    , new Column("性名", String.class)
                    , new Column("年龄", int.class)
                    , new Column("创建时间", Timestamp.class)
                    , new Column("更新", Timestamp.class)
                )
                .writeTo(defaultTestPath);
        }
    }

    private void testStyleProcessor(boolean autoSize) throws SQLException, IOException {
        try (Connection con = getConnection()) {
            new Workbook("statement style processor", author)
                .setAutoSize(autoSize)
                .setConnection(con)
                .addSheet("select id, name, age, create_date, update_date from student"
                    , new Column("学号", int.class)
                    , new Column("性名", String.class)
                    , new Column("年龄", int.class)
                        .setStyleProcessor((o, style, sst) -> {
                            Integer n = (Integer) o;
                            if (n == null || n < 10) {
                                style = Styles.clearFill(style)
                                    | sst.addFill(new Fill(PatternType.solid, Color.orange));
                            }
                            return style;
                        })
                    , new Column("创建时间", Timestamp.class)
                    , new Column("更新", Timestamp.class)
                )
                .writeTo(defaultTestPath);
        }
    }

    private void testIntConversion(boolean autoSize) throws SQLException, IOException {
        try (Connection con = getConnection()) {
            new Workbook("test int conversion statement", author)
                .setConnection(con)
                .setAutoSize(autoSize)
                .watch(Print::println)
                .addSheet("select id, name, age, create_date, update_date from student"
                    , new Column("学号", int.class)
                    , new Column("姓名", String.class)
                    , new Column("年龄", int.class, n -> n > 14 ? "高龄" : n)
                        .setStyleProcessor((o, style, sst) -> {
                            Integer n = (Integer) o;
                            if (n == null || n > 14) {
                                style = Styles.clearFill(style)
                                    | sst.addFill(new Fill(PatternType.solid, Color.orange));
                            }
                            return style;
                        })
                    , new Column("创建时间", Timestamp.class)
                    , new Column("更新", Timestamp.class)
                )
                .writeTo(defaultTestPath);
        }
    }

    @Test public void testConstructor1() throws SQLException, IOException {
        try (Connection con = getConnection()) {
            new Workbook("test statement sheet Constructor1", author)
                .watch(Print::println)
                .addSheet(new StatementSheet(con, "select id, name, age, create_date, update_date from student limit 10"))
                .writeTo(defaultTestPath);
        }
    }

    @Test public void testConstructor2() throws SQLException, IOException {
        try (Connection con = getConnection()) {
            new Workbook("test statement sheet Constructor2", author)
                .watch(Print::println)
                .addSheet(new StatementSheet("Student", con, "select id, name, age, create_date, update_date from student limit 10"))
                .writeTo(defaultTestPath);
        }
    }

    @Test public void testConstructor3() throws SQLException, IOException {
        try (Connection con = getConnection()) {
            new Workbook("test statement sheet Constructor3", author)
                .watch(Print::println)
                .addSheet(new StatementSheet(con, "select id, name, age, create_date, update_date from student where id between ? and ?", ps -> {
                    ps.setInt(1, 10);
                    ps.setInt(2, 20);
                }))
                .writeTo(defaultTestPath);
        }
    }

    @Test public void testConstructor4() throws SQLException, IOException {
        try (Connection con = getConnection()) {
            new Workbook("test statement sheet Constructor4", author)
                .watch(Print::println)
                .addSheet(new StatementSheet("Student", con, "select id, name, age, create_date, update_date from student where id between ? and ?", ps -> {
                    ps.setInt(1, 10);
                    ps.setInt(2, 20);
                }))
                .writeTo(defaultTestPath);
        }
    }

    @Test public void testConstructor5() throws SQLException, IOException {
        try (Connection con = getConnection()) {
            new Workbook("test statement sheet Constructor5", author)
                .watch(Print::println)
                .addSheet(new StatementSheet(con, "select id, name, age, create_date, update_date from student limit 10"
                    , new Column("ID", int.class)
                    , new Column("NAME", String.class)
                    , new Column("AGE", int.class)
                    , new Column("CREATE_DATE", Timestamp.class)
                    , new Column("UPDATE_DATE", Timestamp.class)
                ))
                .writeTo(defaultTestPath);
        }
    }

    @Test public void testConstructor6() throws SQLException, IOException {
        try (Connection con = getConnection()) {
            new Workbook("test statement sheet Constructor6", author)
                .watch(Print::println)
                .addSheet(new StatementSheet("Student", con, "select id, name, age, create_date, update_date from student limit 10"
                    , new Column("ID", int.class)
                    , new Column("NAME", String.class)
                    , new Column("AGE", int.class)
                    , new Column("CREATE_DATE", Timestamp.class)
                    , new Column("UPDATE_DATE", Timestamp.class)
                ))
                .writeTo(defaultTestPath);
        }
    }

    @Test public void testConstructor7() throws SQLException, IOException {
        try (Connection con = getConnection()) {
            new Workbook("test statement sheet Constructor7", author)
                .watch(Print::println)
                .addSheet(new StatementSheet(con, "select id, name, age, create_date, update_date from student where id between ? and ?"
                    , ps -> {
                        ps.setInt(1, 10);
                        ps.setInt(2, 20);
                    }
                    , new Column("ID", int.class)
                    , new Column("NAME", String.class)
                    , new Column("AGE", int.class)
                    , new Column("CREATE_DATE", Timestamp.class)
                    , new Column("UPDATE_DATE", Timestamp.class)
                ))
                .writeTo(defaultTestPath);
        }
    }

    @Test public void testConstructor8() throws SQLException, IOException {
        try (Connection con = getConnection()) {
            new Workbook("test statement sheet Constructor8", author)
                .watch(Print::println)
                .addSheet(new StatementSheet("Student", con, "select id, name, age, create_date, update_date from student where id between ? and ?"
                    , ps -> {
                        ps.setInt(1, 10);
                        ps.setInt(2, 20);
                    }
                    , new Column("ID", int.class)
                    , new Column("NAME", String.class)
                    , new Column("AGE", int.class)
                    , new Column("CREATE_DATE", Timestamp.class)
                    , new Column("UPDATE_DATE", Timestamp.class)
                ))
                .writeTo(defaultTestPath);
        }
    }

    @Test public void testConstructor9() throws IOException {
        try {
            new Workbook("test statement sheet Constructor9", author)
                .watch(Print::println)
                .addSheet(new StatementSheet())
                .writeTo(defaultTestPath);
        } catch (ExcelWriteException e) {
            assert true;
        }
    }

    @Test public void testConstructor10() throws SQLException, IOException {
        try (Connection con = getConnection()) {
            new Workbook("test statement sheet Constructor10", author)
                .watch(Print::println)
                .addSheet(new StatementSheet()
                    .setPs(con.prepareStatement("select id, name, age, create_date, update_date from student limit 10")))
                .writeTo(defaultTestPath);
        }
    }

    @Test public void testConstructor11() throws SQLException, IOException {
        try (Connection con = getConnection()) {
            new Workbook("test statement sheet Constructor11", author)
                .watch(Print::println)
                .addSheet(new StatementSheet("Student")
                    .setPs(con.prepareStatement("select id, name, age, create_date, update_date from student limit 10")))
                .writeTo(defaultTestPath);
        }
    }

    @Test public void testConstructor12() throws SQLException, IOException {
        try (Connection con = getConnection()) {
            new Workbook("test statement sheet Constructor12", author)
                .watch(Print::println)
                .addSheet(new StatementSheet("Student", WaterMark.of(author))
                    .setPs(con.prepareStatement("select id, name, age, create_date, update_date from student limit 10")))
                .writeTo(defaultTestPath);
        }
    }

    @Test public void testConstructor13() throws SQLException, IOException {
        try (Connection con = getConnection()) {
            new Workbook("test statement sheet Constructor13", author)
                .watch(Print::println)
                .addSheet(new StatementSheet("Student", WaterMark.of(author)
                    , new Column("ID", int.class)
                    , new Column("NAME", String.class)
                    , new Column("AGE", int.class))
                    .setPs(con.prepareStatement("select id, name, age, create_date, update_date from student limit 10")))
                .writeTo(defaultTestPath);
        }
    }


    @Test public void testCancelOddStyle() throws SQLException, IOException {
        try (Connection con = getConnection()) {
            new Workbook("test statement sheet cancel odd", author)
                .watch(Print::println)
                .addSheet(new StatementSheet(con, "select id, name, age, create_date, update_date from student limit 10")
                    .setWaterMark(WaterMark.of("TEST"))
                    .cancelOddStyle()
                )
                .writeTo(defaultTestPath);
        }
    }

    @Test public void testDiffTypeFromMetadata() throws SQLException, IOException {
        try (Connection con = getConnection()) {
            new Workbook("test Statement different type from metadata", author)
                .watch(Print::println)
                .addSheet(new StatementSheet(con, "select id, name, age, create_date, update_date from student limit 10"
                    , new Column("ID", String.class)  // Integer in database
                    , new Column("NAME", String.class)
                    , new Column("AGE", String.class) // Integer in database
                    , new Column("CREATE_DATE", String.class) // Timestamp in database
                    , new Column("UPDATE_DATE", String.class) // Timestamp in database
                ))
                .writeTo(defaultTestPath);
        }
    }

    @Test public void testFixWidth() throws SQLException, IOException {
        try (Connection con = getConnection()) {
            new Workbook("test statement fix width", author)
                .watch(Print::println)
                .addSheet(new StatementSheet(con, "select id, name, age, create_date, update_date from student limit 10").fixSize(10))
                .writeTo(defaultTestPath);
        }
    }
}

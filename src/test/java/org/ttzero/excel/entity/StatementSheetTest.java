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
import org.ttzero.excel.entity.style.Fill;
import org.ttzero.excel.entity.style.PatternType;
import org.ttzero.excel.entity.style.Styles;
import org.ttzero.excel.reader.CellType;
import org.ttzero.excel.reader.Col;
import org.ttzero.excel.reader.Drawings;
import org.ttzero.excel.reader.ExcelReader;

import java.awt.Color;
import java.io.IOException;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Timestamp;
import java.util.Iterator;
import java.util.List;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertTrue;

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
            String fileName = ("statement" + (autoSize ? " auto-size" : "")) + ".xlsx"
                , sql = "select id, name, age, create_date, update_date from student order by age";
            new Workbook()
                .setAutoSize(autoSize)
                .addSheet(new StatementSheet(con, sql
                    , new Column("学号", int.class)
                    , new Column("姓名", String.class)
                    , new Column("年龄", int.class)
                    , new Column("创建时间", Timestamp.class).setColIndex(0) // First cell
                    , new Column("更新", Timestamp.class)
                ))
                .writeTo(defaultTestPath.resolve(fileName));

            PreparedStatement ps = con.prepareStatement(sql);
            ResultSet rs = ps.executeQuery();
            try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
                Iterator<org.ttzero.excel.reader.Row> iter = reader.sheet(0).iterator();
                assertTrue(iter.hasNext());
                org.ttzero.excel.reader.Row header = iter.next();
                assertEquals("创建时间", header.getString(0));
                assertEquals("学号", header.getString(1));
                assertEquals("姓名", header.getString(2));
                assertEquals("年龄", header.getString(3));
                assertEquals("更新", header.getString(4));

                while (rs.next()) {
                    assertTrue(iter.hasNext());
                    org.ttzero.excel.reader.Row row = iter.next();

                    // FIXME Timestamp lost millisecond value
                    assertTrue(rs.getTimestamp(4) != null ? rs.getTimestamp(4).getTime() / 1000 == row.getTimestamp(0).getTime() / 1000 : row.getTimestamp(0) == null);
                    assertEquals(rs.getInt(1), (int) row.getInt(1));
                    assertEquals(rs.getString(2), row.getString(2));
                    assertEquals(rs.getInt(3), (int) row.getInt(3));
                    assertTrue(rs.getTimestamp(5) != null ? rs.getTimestamp(5).getTime() / 1000 == row.getTimestamp(4).getTime() / 1000 : row.getTimestamp(4) == null);
                }
            }
            rs.close();
            ps.close();
        }
    }

    private void testStyleProcessor(boolean autoSize) throws SQLException, IOException {
        try (Connection con = getConnection()) {
            String fileName = ("test style processor statement" + (autoSize ? " auto-size" : "")) + ".xlsx"
                , sql = "select id, name, age, create_date, update_date from student order by age";
            new Workbook()
                .setAutoSize(autoSize)
                .addSheet(new StatementSheet(con, sql
                    , new Column("学号", int.class)
                    , new Column("姓名", String.class)
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
                ))
                .writeTo(defaultTestPath.resolve(fileName));

            PreparedStatement ps = con.prepareStatement(sql);
            ResultSet rs = ps.executeQuery();
            try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
                Iterator<org.ttzero.excel.reader.Row> iter = reader.sheet(0).iterator();
                assertTrue(iter.hasNext());
                org.ttzero.excel.reader.Row header = iter.next();
                assertEquals("学号", header.getString(0));
                assertEquals("姓名", header.getString(1));
                assertEquals("年龄", header.getString(2));
                assertEquals("创建时间", header.getString(3));
                assertEquals("更新", header.getString(4));
                while (rs.next()) {
                    assertTrue(iter.hasNext());
                    org.ttzero.excel.reader.Row row = iter.next();

                    assertEquals(rs.getInt(1), (int) row.getInt(0));
                    assertEquals(rs.getString(2), row.getString(1));
                    assertEquals(rs.getInt(3), (int) row.getInt(2));
                    assertTrue(rs.getTimestamp(4) != null ? rs.getTimestamp(4).getTime() / 1000 == row.getTimestamp(3).getTime() / 1000 : row.getTimestamp(0) == null);
                    assertTrue(rs.getTimestamp(5) != null ? rs.getTimestamp(5).getTime() / 1000 == row.getTimestamp(4).getTime() / 1000 : row.getTimestamp(4) == null);

                    Integer age = row.getInt(2);

                    Styles styles = row.getStyles();
                    int style = row.getCellStyle(2);
                    Fill fill = styles.getFill(style);
                    if (age != null && age < 10) {
                        assertTrue(fill != null && fill.getPatternType() == PatternType.solid && fill.getFgColor().equals(Color.orange));
                    } else assertTrue(fill == null || fill.getPatternType() == PatternType.none);
                }
            }
            rs.close();
            ps.close();
        }
    }

    private void testIntConversion(boolean autoSize) throws SQLException, IOException {
        try (Connection con = getConnection()) {
            String fileName = ("test int conversion statement" + (autoSize ? " auto-size" : "")) + ".xlsx"
                , sql = "select id, name, age, create_date, update_date from student";
            new Workbook()
                .setAutoSize(autoSize)
                .addSheet(new StatementSheet(con, sql
                    , new Column("学号", int.class)
                    , new Column("姓名", String.class)
                    , new Column("年龄", int.class, n -> (int) n > 14 ? "高龄" : n)
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
                ))
                .writeTo(defaultTestPath.resolve(fileName));

            PreparedStatement ps = con.prepareStatement(sql);
            ResultSet rs = ps.executeQuery();
            try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
                Iterator<org.ttzero.excel.reader.Row> iter = reader.sheet(0).iterator();
                assertTrue(iter.hasNext());
                org.ttzero.excel.reader.Row header = iter.next();
                assertEquals("学号", header.getString(0));
                assertEquals("姓名", header.getString(1));
                assertEquals("年龄", header.getString(2));
                assertEquals("创建时间", header.getString(3));
                assertEquals("更新", header.getString(4));
                while (rs.next()) {
                    assertTrue(iter.hasNext());
                    org.ttzero.excel.reader.Row row = iter.next();

                    assertEquals(rs.getInt(1), (int) row.getInt(0));
                    assertEquals(rs.getString(2), row.getString(1));
                    assertTrue(rs.getTimestamp(4) != null ? rs.getTimestamp(4).getTime() / 1000 == row.getTimestamp(3).getTime() / 1000 : row.getTimestamp(0) == null);
                    assertTrue(rs.getTimestamp(5) != null ? rs.getTimestamp(5).getTime() / 1000 == row.getTimestamp(4).getTime() / 1000 : row.getTimestamp(4) == null);

                    int age = rs.getInt(3);
                    if (age > 14) assertEquals("高龄", row.getString(2));

                    else assertEquals((int) row.getInt(2), age);
                    Styles styles = row.getStyles();
                    int style = row.getCellStyle(2);
                    Fill fill = styles.getFill(style);
                    if (age > 14) {
                        assertTrue(fill != null && fill.getPatternType() == PatternType.solid && fill.getFgColor().equals(Color.orange));
                    } else assertTrue(fill == null || fill.getPatternType() == PatternType.none);
                }
            }
            rs.close();
            ps.close();
        }
    }

    @Test public void testConstructor1() throws SQLException, IOException {
        try (Connection con = getConnection()) {
            String fileName = "test statement sheet Constructor1.xlsx",
                sql = "select id, name, age, create_date, update_date from student limit 10";
            new Workbook()
                .addSheet(new StatementSheet(con, sql))
                .writeTo(defaultTestPath.resolve(fileName));

            PreparedStatement ps = con.prepareStatement(sql);
            ResultSet rs = ps.executeQuery();
            try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
                Iterator<org.ttzero.excel.reader.Row> iter = reader.sheet(0).iterator();
                assertTrue(iter.hasNext());
                org.ttzero.excel.reader.Row header = iter.next();
                assertEquals("id", header.getString(0));
                assertEquals("name", header.getString(1));
                assertEquals("age", header.getString(2));
                assertEquals("create_date", header.getString(3));
                assertEquals("update_date", header.getString(4));
                while (rs.next()) {
                    assertTrue(iter.hasNext());
                    org.ttzero.excel.reader.Row row = iter.next();

                    assertEquals(rs.getInt(1), (int) row.getInt(0));
                    assertEquals(rs.getString(2), row.getString(1));
                    assertEquals(rs.getInt(3), (int) row.getInt(2));
                    assertTrue(rs.getTimestamp(4) != null ? rs.getTimestamp(4).getTime() / 1000 == row.getTimestamp(3).getTime() / 1000 : row.getTimestamp(0) == null);
                    assertTrue(rs.getTimestamp(5) != null ? rs.getTimestamp(5).getTime() / 1000 == row.getTimestamp(4).getTime() / 1000 : row.getTimestamp(4) == null);
                }
            }
            rs.close();
            ps.close();
        }
    }

    @Test public void testConstructor2() throws SQLException, IOException {
        try (Connection con = getConnection()) {
            String fileName = "test statement sheet Constructor2.xlsx",
                sql = "select id, name, age, create_date, update_date from student limit 10";

            new Workbook()
                .addSheet(new StatementSheet("Student", con, sql))
                .writeTo(defaultTestPath.resolve(fileName));

            PreparedStatement ps = con.prepareStatement(sql);
            ResultSet rs = ps.executeQuery();
            try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
                org.ttzero.excel.reader.Sheet sheet = reader.sheet(0);
                assertEquals("Student", sheet.getName());
                Iterator<org.ttzero.excel.reader.Row> iter = reader.sheet(0).iterator();
                assertTrue(iter.hasNext());
                org.ttzero.excel.reader.Row header = iter.next();
                assertEquals("id", header.getString(0));
                assertEquals("name", header.getString(1));
                assertEquals("age", header.getString(2));
                assertEquals("create_date", header.getString(3));
                assertEquals("update_date", header.getString(4));
                while (rs.next()) {
                    assertTrue(iter.hasNext());
                    org.ttzero.excel.reader.Row row = iter.next();

                    assertEquals(rs.getInt(1), (int) row.getInt(0));
                    assertEquals(rs.getString(2), row.getString(1));
                    assertEquals(rs.getInt(3), (int) row.getInt(2));
                    assertTrue(rs.getTimestamp(4) != null ? rs.getTimestamp(4).getTime() / 1000 == row.getTimestamp(3).getTime() / 1000 : row.getTimestamp(0) == null);
                    assertTrue(rs.getTimestamp(5) != null ? rs.getTimestamp(5).getTime() / 1000 == row.getTimestamp(4).getTime() / 1000 : row.getTimestamp(4) == null);
                }
            }
            rs.close();
            ps.close();
        }
    }

    @Test public void testConstructor3() throws SQLException, IOException {
        try (Connection con = getConnection()) {
            String fileName = "test statement sheet Constructor3.xlsx",
                sql = "select id, name, age, create_date, update_date from student where id between ? and ?";

            new Workbook()
                .addSheet(new StatementSheet(con, sql, ps -> {
                    ps.setInt(1, 10);
                    ps.setInt(2, 20);
                }))
                .writeTo(defaultTestPath.resolve(fileName));

            PreparedStatement ps = con.prepareStatement(sql);
            ps.setInt(1, 10);
            ps.setInt(2, 20);
            ResultSet rs = ps.executeQuery();
            try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
                org.ttzero.excel.reader.Sheet sheet = reader.sheet(0);
                Iterator<org.ttzero.excel.reader.Row> iter = sheet.iterator();
                assertTrue(iter.hasNext());
                org.ttzero.excel.reader.Row header = iter.next();
                assertEquals("id", header.getString(0));
                assertEquals("name", header.getString(1));
                assertEquals("age", header.getString(2));
                assertEquals("create_date", header.getString(3));
                assertEquals("update_date", header.getString(4));
                while (rs.next()) {
                    assertTrue(iter.hasNext());
                    org.ttzero.excel.reader.Row row = iter.next();

                    assertEquals(rs.getInt(1), (int) row.getInt(0));
                    assertEquals(rs.getString(2), row.getString(1));
                    assertEquals(rs.getInt(3), (int) row.getInt(2));
                    assertTrue(rs.getTimestamp(4) != null ? rs.getTimestamp(4).getTime() / 1000 == row.getTimestamp(3).getTime() / 1000 : row.getTimestamp(0) == null);
                    assertTrue(rs.getTimestamp(5) != null ? rs.getTimestamp(5).getTime() / 1000 == row.getTimestamp(4).getTime() / 1000 : row.getTimestamp(4) == null);
                }
            }
            rs.close();
            ps.close();
        }
    }

    @Test public void testConstructor4() throws SQLException, IOException {
        try (Connection con = getConnection()) {
            String fileName = "test statement sheet Constructor4.xlsx",
                sql = "select id, name, age, create_date, update_date from student where id between ? and ?";

            new Workbook()
                .addSheet(new StatementSheet("Student", con, sql, ps -> {
                    ps.setInt(1, 10);
                    ps.setInt(2, 20);
                }))
                .writeTo(defaultTestPath.resolve(fileName));

            PreparedStatement ps = con.prepareStatement(sql);
            ps.setInt(1, 10);
            ps.setInt(2, 20);
            ResultSet rs = ps.executeQuery();
            try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
                org.ttzero.excel.reader.Sheet sheet = reader.sheet(0);
                assertEquals("Student", sheet.getName());
                Iterator<org.ttzero.excel.reader.Row> iter = sheet.iterator();
                assertTrue(iter.hasNext());
                org.ttzero.excel.reader.Row header = iter.next();
                assertEquals("id", header.getString(0));
                assertEquals("name", header.getString(1));
                assertEquals("age", header.getString(2));
                assertEquals("create_date", header.getString(3));
                assertEquals("update_date", header.getString(4));
                while (rs.next()) {
                    assertTrue(iter.hasNext());
                    org.ttzero.excel.reader.Row row = iter.next();

                    assertEquals(rs.getInt(1), (int) row.getInt(0));
                    assertEquals(rs.getString(2), row.getString(1));
                    assertEquals(rs.getInt(3), (int) row.getInt(2));
                    assertTrue(rs.getTimestamp(4) != null ? rs.getTimestamp(4).getTime() / 1000 == row.getTimestamp(3).getTime() / 1000 : row.getTimestamp(0) == null);
                    assertTrue(rs.getTimestamp(5) != null ? rs.getTimestamp(5).getTime() / 1000 == row.getTimestamp(4).getTime() / 1000 : row.getTimestamp(4) == null);
                }
            }
            rs.close();
            ps.close();
        }
    }


    @Test public void testConstructor5() throws SQLException, IOException {
        try (Connection con = getConnection()) {
            String fileName = "test statement sheet Constructor5.xlsx",
                sql = "select id, name, age, create_date, update_date from student limit 10";

            new Workbook()
                .addSheet(new StatementSheet("Student", con, sql
                    , new Column("ID", int.class)
                    , new Column("NAME", String.class)
                    , new Column("AGE", int.class)
                    , new Column("CREATE_DATE", Timestamp.class)
                    , new Column("UPDATE_DATE", Timestamp.class)
                ))
                .writeTo(defaultTestPath.resolve(fileName));

            PreparedStatement ps = con.prepareStatement(sql);
            ResultSet rs = ps.executeQuery();
            try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
                org.ttzero.excel.reader.Sheet sheet = reader.sheet(0);
                assertEquals("Student", sheet.getName());
                Iterator<org.ttzero.excel.reader.Row> iter = sheet.iterator();
                // assert header row
                assertTrue(iter.hasNext());
                org.ttzero.excel.reader.Row header = iter.next();
                assertEquals("ID", header.getString(0));
                assertEquals("NAME", header.getString(1));
                assertEquals("AGE", header.getString(2));
                assertEquals("CREATE_DATE", header.getString(3));
                assertEquals("UPDATE_DATE", header.getString(4));

                while (rs.next()) {
                    assertTrue(iter.hasNext());
                    org.ttzero.excel.reader.Row row = iter.next();

                    assertEquals(rs.getInt(1), (int) row.getInt(0));
                    assertEquals(rs.getString(2), row.getString(1));
                    assertEquals(rs.getInt(3), (int) row.getInt(2));
                    assertTrue(rs.getTimestamp(4) != null ? rs.getTimestamp(4).getTime() / 1000 == row.getTimestamp(3).getTime() / 1000 : row.getTimestamp(0) == null);
                    assertTrue(rs.getTimestamp(5) != null ? rs.getTimestamp(5).getTime() / 1000 == row.getTimestamp(4).getTime() / 1000 : row.getTimestamp(4) == null);
                }
            }
            rs.close();
            ps.close();
        }
    }

    @Test public void testConstructor6() throws SQLException, IOException {
        try (Connection con = getConnection()) {
            String fileName = "test statement sheet Constructor6.xlsx",
                sql = "select id, name, age, create_date, update_date from student where id between ? and ?";

            new Workbook()
                .addSheet(new StatementSheet(con, sql
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
                .writeTo(defaultTestPath.resolve(fileName));

            PreparedStatement ps = con.prepareStatement(sql);
            ps.setInt(1, 10);
            ps.setInt(2, 20);
            ResultSet rs = ps.executeQuery();
            try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
                org.ttzero.excel.reader.Sheet sheet = reader.sheet(0);
                Iterator<org.ttzero.excel.reader.Row> iter = sheet.iterator();
                // assert header row
                assertTrue(iter.hasNext());
                org.ttzero.excel.reader.Row header = iter.next();
                assertEquals("ID", header.getString(0));
                assertEquals("NAME", header.getString(1));
                assertEquals("AGE", header.getString(2));
                assertEquals("CREATE_DATE", header.getString(3));
                assertEquals("UPDATE_DATE", header.getString(4));

                while (rs.next()) {
                    assertTrue(iter.hasNext());
                    org.ttzero.excel.reader.Row row = iter.next();

                    assertEquals(rs.getInt(1), (int) row.getInt(0));
                    assertEquals(rs.getString(2), row.getString(1));
                    assertEquals(rs.getInt(3), (int) row.getInt(2));
                    assertTrue(rs.getTimestamp(4) != null ? rs.getTimestamp(4).getTime() / 1000 == row.getTimestamp(3).getTime() / 1000 : row.getTimestamp(0) == null);
                    assertTrue(rs.getTimestamp(5) != null ? rs.getTimestamp(5).getTime() / 1000 == row.getTimestamp(4).getTime() / 1000 : row.getTimestamp(4) == null);
                }
            }
            rs.close();
            ps.close();
        }
    }


    @Test(expected = ExcelWriteException.class) public void testConstructor9() throws IOException {
        new Workbook("test statement sheet Constructor9")
            .addSheet(new StatementSheet())
            .writeTo(defaultTestPath);
    }

    @Test public void testConstructor10() throws SQLException, IOException {
        try (Connection con = getConnection()) {
            String fileName = "test statement sheet Constructor10.xlsx",
                sql = "select id, name, age, create_date, update_date from student limit 10";

            new Workbook()
                .addSheet(new StatementSheet()
                    .setStatement(con.prepareStatement(sql)))
                .writeTo(defaultTestPath.resolve(fileName));

            PreparedStatement ps = con.prepareStatement(sql);
            ResultSet rs = ps.executeQuery();
            try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
                org.ttzero.excel.reader.Sheet sheet = reader.sheet(0);
                Iterator<org.ttzero.excel.reader.Row> iter = sheet.iterator();
                assertTrue(iter.hasNext());
                org.ttzero.excel.reader.Row header = iter.next();
                assertEquals("id", header.getString(0));
                assertEquals("name", header.getString(1));
                assertEquals("age", header.getString(2));
                assertEquals("create_date", header.getString(3));
                assertEquals("update_date", header.getString(4));
                while (rs.next()) {
                    assertTrue(iter.hasNext());
                    org.ttzero.excel.reader.Row row = iter.next();

                    assertEquals(rs.getInt(1), (int) row.getInt(0));
                    assertEquals(rs.getString(2), row.getString(1));
                    assertEquals(rs.getInt(3), (int) row.getInt(2));
                    assertTrue(rs.getTimestamp(4) != null ? rs.getTimestamp(4).getTime() / 1000 == row.getTimestamp(3).getTime() / 1000 : row.getTimestamp(0) == null);
                    assertTrue(rs.getTimestamp(5) != null ? rs.getTimestamp(5).getTime() / 1000 == row.getTimestamp(4).getTime() / 1000 : row.getTimestamp(4) == null);
                }
            }
            rs.close();
            ps.close();
        }
    }

    @Test public void testConstructor11() throws SQLException, IOException {
        try (Connection con = getConnection()) {
            String fileName = "test statement sheet Constructor11.xlsx",
                sql = "select id, name, age, create_date, update_date from student limit 10";

            new Workbook()
                .addSheet(new StatementSheet("Student")
                    .setStatement(con.prepareStatement(sql)))
                .writeTo(defaultTestPath.resolve(fileName));

            PreparedStatement ps = con.prepareStatement(sql);
            ResultSet rs = ps.executeQuery();
            try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
                org.ttzero.excel.reader.Sheet sheet = reader.sheet(0);
                assertEquals("Student", sheet.getName());
                Iterator<org.ttzero.excel.reader.Row> iter = sheet.iterator();
                assertTrue(iter.hasNext());
                // Header row
                org.ttzero.excel.reader.Row header = iter.next();
                assertEquals("id", header.getString(0));
                assertEquals("name", header.getString(1));
                assertEquals("age", header.getString(2));
                assertEquals("create_date", header.getString(3));
                assertEquals("update_date", header.getString(4));
                // Body rows
                while (rs.next()) {
                    assertTrue(iter.hasNext());
                    org.ttzero.excel.reader.Row row = iter.next();

                    assertEquals(rs.getInt(1), (int) row.getInt(0));
                    assertEquals(rs.getString(2), row.getString(1));
                    assertEquals(rs.getInt(3), (int) row.getInt(2));
                    assertTrue(rs.getTimestamp(4) != null ? rs.getTimestamp(4).getTime() / 1000 == row.getTimestamp(3).getTime() / 1000 : row.getTimestamp(0) == null);
                    assertTrue(rs.getTimestamp(5) != null ? rs.getTimestamp(5).getTime() / 1000 == row.getTimestamp(4).getTime() / 1000 : row.getTimestamp(4) == null);
                }
            }
            rs.close();
            ps.close();
        }
    }

    @Test public void testConstructor12() throws SQLException, IOException {
        try (Connection con = getConnection()) {
            String fileName = "test statement sheet Constructor12.xlsx",
                sql = "select id, name, age, create_date, update_date from student limit 10";

            new Workbook()
                .addSheet(new StatementSheet("Student", WaterMark.of(author))
                    .setStatement(con.prepareStatement(sql)))
                .writeTo(defaultTestPath.resolve(fileName));

            PreparedStatement ps = con.prepareStatement(sql);
            ResultSet rs = ps.executeQuery();
            try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
                org.ttzero.excel.reader.Sheet sheet = reader.sheet(0);
                assertEquals("Student", sheet.getName());
                Iterator<org.ttzero.excel.reader.Row> iter = sheet.iterator();
                assertTrue(iter.hasNext());
                // Header row
                org.ttzero.excel.reader.Row header = iter.next();
                assertEquals("id", header.getString(0));
                assertEquals("name", header.getString(1));
                assertEquals("age", header.getString(2));
                assertEquals("create_date", header.getString(3));
                assertEquals("update_date", header.getString(4));
                // Body rows
                while (rs.next()) {
                    assertTrue(iter.hasNext());
                    org.ttzero.excel.reader.Row row = iter.next();

                    assertEquals(rs.getInt(1), (int) row.getInt(0));
                    assertEquals(rs.getString(2), row.getString(1));
                    assertEquals(rs.getInt(3), (int) row.getInt(2));
                    assertTrue(rs.getTimestamp(4) != null ? rs.getTimestamp(4).getTime() / 1000 == row.getTimestamp(3).getTime() / 1000 : row.getTimestamp(0) == null);
                    assertTrue(rs.getTimestamp(5) != null ? rs.getTimestamp(5).getTime() / 1000 == row.getTimestamp(4).getTime() / 1000 : row.getTimestamp(4) == null);
                }

                // Water Mark
                List<Drawings.Picture> pictures = sheet.listPictures();
                assertEquals(pictures.size(), 1);
                assertTrue(pictures.get(0).isBackground());
            }
            rs.close();
            ps.close();
        }
    }

    @Test public void testConstructor13() throws SQLException, IOException {
        try (Connection con = getConnection()) {
            String fileName = "test statement sheet Constructor13.xlsx",
                sql = "select id, name, age, create_date, update_date from student limit 10";

            new Workbook()
                .addSheet(new StatementSheet("Student", WaterMark.of(author)
                    , new Column("ID", int.class)
                    , new Column("NAME", String.class)
                    , new Column("AGE", int.class))
                    .setStatement(con.prepareStatement(sql)))
                .writeTo(defaultTestPath.resolve(fileName).toFile());

            PreparedStatement ps = con.prepareStatement(sql);
            ResultSet rs = ps.executeQuery();
            try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
                org.ttzero.excel.reader.Sheet sheet = reader.sheet(0);
                assertEquals("Student", sheet.getName());
                Iterator<org.ttzero.excel.reader.Row> iter = sheet.iterator();
                // assert header row
                assertTrue(iter.hasNext());
                org.ttzero.excel.reader.Row header = iter.next();
                assertEquals("ID", header.getString(0));
                assertEquals("NAME", header.getString(1));
                assertEquals("AGE", header.getString(2));

                while (rs.next()) {
                    assertTrue(iter.hasNext());
                    org.ttzero.excel.reader.Row row = iter.next();

                    assertEquals(rs.getInt(1), (int) row.getInt(0));
                    assertEquals(rs.getString(2), row.getString(1));
                    assertEquals(rs.getInt(3), (int) row.getInt(2));
                }
            }
            rs.close();
            ps.close();
        }
    }


    @Test public void testCancelOddStyle() throws SQLException, IOException {
        try (Connection con = getConnection()) {
            String fileName = "test statement sheet cancel odd.xlsx",
                sql = "select id, name, age, create_date, update_date from student limit 10";

            new Workbook()
                .addSheet(new StatementSheet(con, sql)
                    .setWaterMark(WaterMark.of("TEST"))
                    .cancelZebraLine()
                )
                .writeTo(defaultTestPath.resolve(fileName));

            PreparedStatement ps = con.prepareStatement(sql);
            ResultSet rs = ps.executeQuery();
            try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
                org.ttzero.excel.reader.Sheet sheet = reader.sheet(0);
                Iterator<org.ttzero.excel.reader.Row> iter = sheet.iterator();
                assertTrue(iter.hasNext());
                // Header row
                org.ttzero.excel.reader.Row header = iter.next();
                assertEquals("id", header.getString(0));
                assertEquals("name", header.getString(1));
                assertEquals("age", header.getString(2));
                assertEquals("create_date", header.getString(3));
                assertEquals("update_date", header.getString(4));
                // Body rows
                while (rs.next()) {
                    assertTrue(iter.hasNext());
                    org.ttzero.excel.reader.Row row = iter.next();

                    assertEquals(rs.getInt(1), (int) row.getInt(0));
                    assertEquals(rs.getString(2), row.getString(1));
                    assertEquals(rs.getInt(3), (int) row.getInt(2));
                    assertTrue(rs.getTimestamp(4) != null ? rs.getTimestamp(4).getTime() / 1000 == row.getTimestamp(3).getTime() / 1000 : row.getTimestamp(0) == null);
                    assertTrue(rs.getTimestamp(5) != null ? rs.getTimestamp(5).getTime() / 1000 == row.getTimestamp(4).getTime() / 1000 : row.getTimestamp(4) == null);

                    Styles styles = row.getStyles();
                    int style = row.getCellStyle(0);
                    Fill fill = styles.getFill(style);
                    assertTrue(fill == null || fill.getPatternType() == PatternType.none);
                }

                // Water Mark
                List<Drawings.Picture> pictures = sheet.listPictures();
                assertEquals(pictures.size(), 1);
                assertTrue(pictures.get(0).isBackground());
            }
            rs.close();
            ps.close();
        }
    }

    @Test public void testDiffTypeFromMetadata() throws SQLException, IOException {
        try (Connection con = getConnection()) {
            String fileName = "test Statement different type from metadata.xlsx",
                sql = "select id, name, age, create_date, update_date from student limit 10";

            new Workbook()
                .addSheet(new StatementSheet(con, sql
                    , new Column("ID", String.class)  // Integer in database
                    , new Column("NAME", String.class)
                    , new Column("AGE", String.class) // Integer in database
                    , new Column("CREATE_DATE", String.class) // Timestamp in database
                    , new Column("UPDATE_DATE", String.class) // Timestamp in database
                ))
                .writeTo(defaultTestPath.resolve(fileName));

            PreparedStatement ps = con.prepareStatement(sql);
            ResultSet rs = ps.executeQuery();
            try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
                org.ttzero.excel.reader.Sheet sheet = reader.sheet(0);
                Iterator<org.ttzero.excel.reader.Row> iter = sheet.iterator();
                assertTrue(iter.hasNext());
                // Header row
                org.ttzero.excel.reader.Row header = iter.next();
                assertEquals("ID", header.getString(0));
                assertEquals("NAME", header.getString(1));
                assertEquals("AGE", header.getString(2));
                assertEquals("CREATE_DATE", header.getString(3));
                assertEquals("UPDATE_DATE", header.getString(4));
                // Body rows
                while (rs.next()) {
                    assertTrue(iter.hasNext());
                    org.ttzero.excel.reader.Row row = iter.next();

                    assertEquals(row.getCellType(0), CellType.STRING);
                    assertEquals(row.getCellType(1), CellType.STRING);
                    assertEquals(row.getCellType(2), CellType.STRING);
                    assertEquals(row.getCellType(3), CellType.STRING);
                    assertTrue(row.getCellType(4) == CellType.STRING || row.getCellType(4) == CellType.BLANK);

                    assertEquals(rs.getInt(1), (int) row.getInt(0));
                    assertEquals(rs.getString(2), row.getString(1));
                    assertEquals(rs.getInt(3), (int) row.getInt(2));
                    assertTrue(rs.getTimestamp(4) != null ? rs.getTimestamp(4).getTime() / 1000 == row.getTimestamp(3).getTime() / 1000 : row.getTimestamp(0) == null);
                    assertTrue(rs.getTimestamp(5) != null ? rs.getTimestamp(5).getTime() / 1000 == row.getTimestamp(4).getTime() / 1000 : row.getTimestamp(4) == null);
                }
            }
            rs.close();
            ps.close();
        }
    }

    @Test public void testFixWidth() throws SQLException, IOException {
        try (Connection con = getConnection()) {
            String fileName = "test statement fix width.xlsx",
                sql = "select id, name, age, create_date, update_date from student limit 10";

            new Workbook()
                .addSheet(new StatementSheet(con, sql).fixedSize(10))
                .writeTo(defaultTestPath.resolve(fileName));

            PreparedStatement ps = con.prepareStatement(sql);
            ResultSet rs = ps.executeQuery();
            try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
                org.ttzero.excel.reader.FullSheet sheet = reader.sheet(0).asFullSheet();
                Iterator<org.ttzero.excel.reader.Row> iter = sheet.iterator();
                assertTrue(iter.hasNext());
                // Header row
                org.ttzero.excel.reader.Row header = iter.next();
                assertEquals("id", header.getString(0));
                assertEquals("name", header.getString(1));
                assertEquals("age", header.getString(2));
                assertEquals("create_date", header.getString(3));
                assertEquals("update_date", header.getString(4));
                // Body rows
                while (rs.next()) {
                    assertTrue(iter.hasNext());
                    org.ttzero.excel.reader.Row row = iter.next();

                    assertEquals(rs.getInt(1), (int) row.getInt(0));
                    assertEquals(rs.getString(2), row.getString(1));
                    assertEquals(rs.getInt(3), (int) row.getInt(2));
                    assertTrue(rs.getTimestamp(4) != null ? rs.getTimestamp(4).getTime() / 1000 == row.getTimestamp(3).getTime() / 1000 : row.getTimestamp(0) == null);
                    assertTrue(rs.getTimestamp(5) != null ? rs.getTimestamp(5).getTime() / 1000 == row.getTimestamp(4).getTime() / 1000 : row.getTimestamp(4) == null);

                    // assert column width equals 10
                    List<Col> cols = sheet.getCols();
                    assertNotNull(cols);
                    assertEquals(cols.size(), 5);
                    for (Col col : cols) assertTrue(col.width - 10.0D <= 0.8D);
                }
            }
            rs.close();
            ps.close();
        }
    }
}

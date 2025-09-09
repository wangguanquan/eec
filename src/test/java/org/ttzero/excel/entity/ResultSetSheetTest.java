/*
 * Copyright (c) 2017-2019, guanquan.wang@hotmail.com All Rights Reserved.
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
import org.ttzero.excel.reader.Drawings;
import org.ttzero.excel.reader.ExcelReader;
import org.ttzero.excel.reader.Row;

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
 * @author guanquan.wang at 2019-04-28 21:50
 */
public class ResultSetSheetTest extends SQLWorkbookTest {
    @Test public void testWrite() throws SQLException, IOException {
        String fileName = "result set.xlsx",
            sql = "select id, name, age, create_date, update_date from student limit 10";

        try (
            Connection con = getConnection();
            PreparedStatement ps = con.prepareStatement(sql);
            ResultSet rs = ps.executeQuery()
        ) {
            new Workbook()
                .addSheet(new ResultSetSheet(new Column("学号", int.class)
                    , new Column("姓名", String.class)
                    , new Column("年龄", Integer.class)
                    , new Column("创建时间", Timestamp.class)
                    , new Column("更新", Timestamp.class)
                ).setResultSet(rs))
                .writeTo(defaultTestPath.resolve(fileName));

            PreparedStatement ps1 = con.prepareStatement(sql);
            ResultSet rs1 = ps1.executeQuery();
            try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
                Iterator<Row> iter = reader.sheet(0).iterator();
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
                }
            }
            rs1.close();
            ps1.close();
        }
    }

    @Test public void testStyleDesign4RS() throws IOException, SQLException {
        String fileName = "test global style design for ResultSet.xlsx",
            sql = "select id, name, age, create_date, update_date from student limit 10";
        try (
            Connection con = getConnection();
            PreparedStatement ps = con.prepareStatement(sql);
            ResultSet rs = ps.executeQuery()
        ) {
            new Workbook()
                .addSheet(new ResultSetSheet().setResultSet(rs).setStyleProcessor((rst, style, sst) -> {
                    try {
                        if (rst.getInt("age") > 14)
                            style = sst.modifyFill(style, new Fill(PatternType.solid, Color.yellow));
                    } catch (SQLException ex) {
                        // Ignore
                    }
                    return style;
                }))
                .writeTo(defaultTestPath.resolve(fileName));

            PreparedStatement ps1 = con.prepareStatement(sql);
            ResultSet rs1 = ps1.executeQuery();
            try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
                Iterator<Row> iter = reader.sheet(0).iterator();
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

                    Styles styles = row.getStyles();
                    int style = row.getCellStyle(2);
                    Fill fill = styles.getFill(style);
                    if (rs.getInt(3) > 14) {
                        assertTrue(fill != null && fill.getPatternType() == PatternType.solid && fill.getFgColor().equals(Color.yellow));
                    } else assertTrue( fill == null || fill.getPatternType() == PatternType.none);
                }
            }
            rs1.close();
            ps1.close();
        }
    }

    @Test(expected = ExcelWriteException.class) public void testConstructor1() throws IOException {
        new Workbook("test ResultSet sheet Constructor1", author)
                .addSheet(new ResultSetSheet())
                .writeTo(defaultTestPath);
    }

    @Test public void testConstructor2() throws SQLException, IOException {
        String fileName = "test ResultSet sheet Constructor2.xlsx",
            sql = "select id, name, age, create_date, update_date from student limit 10";
        try (
            Connection con = getConnection();
            PreparedStatement ps = con.prepareStatement(sql);
            ResultSet rs = ps.executeQuery()
        ) {
            new Workbook()
                .addSheet(new ResultSetSheet().setResultSet(rs))
                .writeTo(defaultTestPath.resolve(fileName));

            PreparedStatement ps1 = con.prepareStatement(sql);
            ResultSet rs1 = ps1.executeQuery();
            try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
                Iterator<Row> iter = reader.sheet(0).iterator();
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
            rs1.close();
            ps1.close();
        }
    }

    @Test public void testConstructor3() throws SQLException, IOException {
        String fileName = "test ResultSet sheet Constructor3.xlsx",
            sql = "select id, name, age, create_date, update_date from student limit 10";
        try (
            Connection con = getConnection();
            PreparedStatement ps = con.prepareStatement(sql);
            ResultSet rs = ps.executeQuery()
        ) {
            new Workbook()
                .addSheet(new ResultSetSheet("Student").setResultSet(rs))
                .writeTo(defaultTestPath.resolve(fileName));

            PreparedStatement ps1 = con.prepareStatement(sql);
            ResultSet rs1 = ps1.executeQuery();
            try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
                org.ttzero.excel.reader.Sheet sheet = reader.sheet(0);
                assertEquals("Student", sheet.getName());
                Iterator<Row> iter = sheet.iterator();
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
            rs1.close();
            ps1.close();
        }
    }

    @Test public void testConstructor4() throws SQLException, IOException {
        String fileName = "test ResultSet sheet Constructor4.xlsx",
            sql = "select id, name, age, create_date, update_date from student limit 10";
        try (
            Connection con = getConnection();
            PreparedStatement ps = con.prepareStatement(sql);
            ResultSet rs = ps.executeQuery()
        ) {
            new Workbook()
                .addSheet(new ResultSetSheet("Student"
                    , new Column("ID", int.class)
                    , new Column("NAME", String.class)
                    , new Column("AGE", int.class)
                    , new Column("CREATE_DATE", Timestamp.class)
                    , new Column("UPDATE_DATE", Timestamp.class)
                ).setResultSet(rs))
                .writeTo(defaultTestPath.resolve(fileName));

            PreparedStatement ps1 = con.prepareStatement(sql);
            ResultSet rs1 = ps1.executeQuery();
            try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
                org.ttzero.excel.reader.Sheet sheet = reader.sheet(0);
                assertEquals("Student", sheet.getName());
                Iterator<Row> iter = sheet.iterator();
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
            rs1.close();
            ps1.close();
        }
    }

    @Test public void testConstructor5() throws SQLException, IOException {
        String fileName = "test ResultSet sheet Constructor5.xlsx",
            sql = "select id, name, age, create_date, update_date from student limit 10";
        try (
            Connection con = getConnection();
            PreparedStatement ps = con.prepareStatement(sql);
            ResultSet rs = ps.executeQuery()
        ) {
            new Workbook()
                .addSheet(new ResultSetSheet("Student"
                    , new Column("ID", int.class)
                    , new Column("NAME", String.class)
                    , new Column("AGE", int.class)
                    , new Column("CREATE_DATE", Timestamp.class)
                    , new Column("UPDATE_DATE", Timestamp.class)
                ).setResultSet(rs).setWatermark(Watermark.of(author)))
                .writeTo(defaultTestPath.resolve(fileName));

            PreparedStatement ps1 = con.prepareStatement(sql);
            ResultSet rs1 = ps1.executeQuery();
            try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
                org.ttzero.excel.reader.Sheet sheet = reader.sheet(0);
                assertEquals("Student", sheet.getName());
                Iterator<Row> iter = sheet.iterator();
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

                // Water Mark
                List<Drawings.Picture> pictures = sheet.listPictures();
                assertEquals(pictures.size(), 1);
                assertTrue(pictures.get(0).isBackground());
            }
            rs1.close();
            ps1.close();
        }
    }

    @Test public void testConstructor6() throws SQLException, IOException {
        String fileName = "test ResultSet sheet Constructor6.xlsx",
            sql = "select id, name, age, create_date, update_date from student limit 10";
        try (
            Connection con = getConnection();
            PreparedStatement ps = con.prepareStatement(sql);
            ResultSet rs = ps.executeQuery()
        ) {
            new Workbook()
                .addSheet(new ResultSetSheet(rs))
                .writeTo(defaultTestPath.resolve(fileName));

            PreparedStatement ps1 = con.prepareStatement(sql);
            ResultSet rs1 = ps1.executeQuery();
            try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
                org.ttzero.excel.reader.Sheet sheet = reader.sheet(0);
                Iterator<Row> iter = sheet.iterator();
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
            rs1.close();
            ps1.close();
        }
    }

    @Test public void testConstructor7() throws SQLException, IOException {
        String fileName = "test ResultSet sheet Constructor7.xlsx",
            sql = "select id, name, age, create_date, update_date from student limit 10";
        try (
            Connection con = getConnection();
            PreparedStatement ps = con.prepareStatement(sql);
            ResultSet rs = ps.executeQuery()
        ) {
            new Workbook()
                .addSheet(new ResultSetSheet("Student", rs))
                .writeTo(defaultTestPath.resolve(fileName));

            PreparedStatement ps1 = con.prepareStatement(sql);
            ResultSet rs1 = ps1.executeQuery();
            try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
                org.ttzero.excel.reader.Sheet sheet = reader.sheet(0);
                assertEquals("Student", sheet.getName());
                Iterator<Row> iter = sheet.iterator();
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
            rs1.close();
            ps1.close();
        }
    }

    @Test public void testConstructor8() throws SQLException, IOException {
        String fileName = "test ResultSet sheet Constructor8.xlsx",
            sql = "select id, name, age, create_date, update_date from student limit 10";
        try (
            Connection con = getConnection();
            PreparedStatement ps = con.prepareStatement(sql);
            ResultSet rs = ps.executeQuery()
        ) {
            new Workbook()
                .addSheet(new ResultSetSheet(rs
                    , new Column("ID", int.class)
                    , new Column("NAME", String.class)
                    , new Column("AGE", int.class)
                    , new Column("CREATE_DATE", Timestamp.class)
                    , new Column("UPDATE_DATE", Timestamp.class)
                ))
                .writeTo(defaultTestPath.resolve(fileName));

            PreparedStatement ps1 = con.prepareStatement(sql);
            ResultSet rs1 = ps1.executeQuery();
            try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
                org.ttzero.excel.reader.Sheet sheet = reader.sheet(0);
                Iterator<Row> iter = sheet.iterator();
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
            rs1.close();
            ps1.close();
        }
    }

    @Test public void testConstructor9() throws SQLException, IOException {
        String fileName = "test ResultSet sheet Constructor9.xlsx",
            sql = "select id, name, age, create_date, update_date from student limit 10";
        try (
            Connection con = getConnection();
            PreparedStatement ps = con.prepareStatement(sql);
            ResultSet rs = ps.executeQuery()
        ) {
            new Workbook()
                .addSheet(new ResultSetSheet("Student", rs
                    , new Column("ID", int.class)
                    , new Column("NAME", String.class)
                    , new Column("AGE", int.class)
                    , new Column("CREATE_DATE", Timestamp.class)
                    , new Column("UPDATE_DATE", Timestamp.class)
                ))
                .writeTo(defaultTestPath.resolve(fileName));

            PreparedStatement ps1 = con.prepareStatement(sql);
            ResultSet rs1 = ps1.executeQuery();
            try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
                org.ttzero.excel.reader.Sheet sheet = reader.sheet(0);
                assertEquals("Student", sheet.getName());
                Iterator<Row> iter = sheet.iterator();
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
            rs1.close();
            ps1.close();
        }
    }

    @Test public void testConstructor10() throws SQLException, IOException {
        String fileName = "test ResultSet sheet Constructor10.xlsx",
            sql = "select id, name, age, create_date, update_date from student limit 10";
        try (
            Connection con = getConnection();
            PreparedStatement ps = con.prepareStatement(sql);
            ResultSet rs = ps.executeQuery()
        ) {
            new Workbook()
                .addSheet(new ResultSetSheet(rs
                    , new Column("ID", int.class)
                    , new Column("NAME", String.class)
                    , new Column("AGE", int.class)
                    , new Column("CREATE_DATE", Timestamp.class)
                    , new Column("UPDATE_DATE", Timestamp.class)
                ).setWatermark(Watermark.of(author)))
                .writeTo(defaultTestPath.resolve(fileName));

            PreparedStatement ps1 = con.prepareStatement(sql);
            ResultSet rs1 = ps1.executeQuery();
            try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
                org.ttzero.excel.reader.Sheet sheet = reader.sheet(0);
                Iterator<Row> iter = sheet.iterator();
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

                // Water Mark
                List<Drawings.Picture> pictures = sheet.listPictures();
                assertEquals(pictures.size(), 1);
                assertTrue(pictures.get(0).isBackground());
            }
            rs1.close();
            ps1.close();
        }
    }

    @Test public void testConstructor11() throws SQLException, IOException {
        String fileName = "test ResultSet sheet Constructor11.xlsx",
            sql = "select id, name, age, create_date, update_date from student limit 10";
        try (
            Connection con = getConnection();
            PreparedStatement ps = con.prepareStatement(sql);
            ResultSet rs = ps.executeQuery()
        ) {
            new Workbook()
                .addSheet(new ResultSetSheet("Student", rs
                    , new Column("ID", int.class)
                    , new Column("NAME", String.class)
                    , new Column("AGE", int.class)
                    , new Column("CREATE_DATE", Timestamp.class)
                    , new Column("UPDATE_DATE", Timestamp.class)
                )).setWatermark(Watermark.of(author))
                .writeTo(defaultTestPath.resolve(fileName));

            PreparedStatement ps1 = con.prepareStatement(sql);
            ResultSet rs1 = ps1.executeQuery();
            try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
                org.ttzero.excel.reader.Sheet sheet = reader.sheet(0);
                Iterator<Row> iter = sheet.iterator();
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

                // Water Mark
                List<Drawings.Picture> pictures = sheet.listPictures();
                assertEquals(pictures.size(), 1);
                assertTrue(pictures.get(0).isBackground());
            }
            rs1.close();
            ps1.close();
        }
    }

    @Test public void testDiffTypeFromMetadata() throws SQLException, IOException {
        String fileName = "test ResultSet different type from metadata.xlsx",
            sql = "select id, name, age, create_date, update_date from student limit 10";
        try (
            Connection con = getConnection();
            PreparedStatement ps = con.prepareStatement(sql);
            ResultSet rs = ps.executeQuery()
        ) {
            new Workbook("test ResultSet different type from metadata", author)
                .addSheet(new ResultSetSheet("Student", rs
                    , new Column("ID", String.class)  // Integer in database
                    , new Column("NAME", String.class)
                    , new Column("AGE", String.class) // // Integer in database
                    , new Column("CREATE_DATE", String.class)
                    , new Column("UPDATE_DATE", String.class)
                )).setWatermark(Watermark.of(author))
                .writeTo(defaultTestPath.resolve(fileName));

            PreparedStatement ps1 = con.prepareStatement(sql);
            ResultSet rs1 = ps1.executeQuery();
            try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
                org.ttzero.excel.reader.Sheet sheet = reader.sheet(0);
                Iterator<Row> iter = sheet.iterator();
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

                // Water Mark
                List<Drawings.Picture> pictures = sheet.listPictures();
                assertEquals(pictures.size(), 1);
                assertTrue(pictures.get(0).isBackground());
            }
            rs1.close();
            ps1.close();
        }
    }

    @Test public void testTypes() throws SQLException, IOException {
        final String fileName = "test types.xlsx";
        try (
            Connection con = getConnection();
            PreparedStatement ps = con.prepareStatement("select * from types_test");
            ResultSet rs = ps.executeQuery()
        ) {
            new Workbook()
                .addSheet(new ResultSetSheet(rs))
                .writeTo(defaultTestPath.resolve(fileName));

            PreparedStatement ps1 = con.prepareStatement("select * from types_test");
            ResultSet rs1 = ps1.executeQuery();
            try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
                org.ttzero.excel.reader.Sheet sheet = reader.sheet(0);
                Iterator<Row> iter = sheet.iterator();
                assertTrue(iter.hasNext());
                org.ttzero.excel.reader.Row header = iter.next();
                assertEquals("id", header.getString(0));
                assertEquals("t_bit", header.getString(1));
                assertEquals("t_tinyint", header.getString(2));
                assertEquals("t_smallint", header.getString(3));
                assertEquals("t_int", header.getString(4));
                assertEquals("t_bigint", header.getString(5));
                assertEquals("t_float", header.getString(6));
                assertEquals("t_double", header.getString(7));
                assertEquals("t_varchar", header.getString(8));
                assertEquals("t_char", header.getString(9));
                assertEquals("t_date", header.getString(10));
                assertEquals("t_datetime", header.getString(11));
                assertEquals("t_timestamp", header.getString(12));

                assertTrue(iter.hasNext());
                assertTrue(rs1.next());
                // Row1
                org.ttzero.excel.reader.Row row = iter.next();

                assertEquals(row.getCellType(0), CellType.INTEGER);
                assertEquals(row.getCellType(1), "mysql".equals(protocol) ? CellType.BOOLEAN : CellType.INTEGER);
                assertEquals(row.getCellType(2), CellType.INTEGER);
                assertEquals(row.getCellType(3), CellType.INTEGER);
                assertEquals(row.getCellType(4), CellType.INTEGER);
                assertEquals(row.getCellType(5), CellType.LONG);
                assertEquals(row.getCellType(6), CellType.DECIMAL);
                assertEquals(row.getCellType(7), CellType.DECIMAL);
                assertEquals(row.getCellType(8), CellType.STRING);
                assertEquals(row.getCellType(9), CellType.STRING);
                assertEquals(row.getCellType(10), CellType.DATE);
                assertEquals(row.getCellType(11), CellType.DATE);
                assertEquals(row.getCellType(12), CellType.DATE);

                assertEquals(rs1.getInt(1), (int) row.getInt(0));
                if ("mysql".equals(protocol)) {
                    assertEquals(rs1.getBoolean(2), row.getBoolean(1));
                } else {
                    assertEquals(rs1.getInt(2), (int) row.getInt(1));
                }
                assertEquals(rs1.getInt(3), (int) row.getInt(2));
                assertEquals(rs1.getInt(4), (int) row.getInt(3));
                assertEquals(rs1.getInt(5), (int) row.getInt(4));
                assertEquals(rs1.getLong(6), (long) row.getLong(5));
                assertTrue(rs1.getDouble(7) - row.getDouble(6) <= 0.0001);
                assertTrue(rs1.getDouble(8) - row.getDouble(7) <= 0.0001);
                assertEquals(rs1.getString(9), row.getString(8));
                assertEquals(rs1.getString(10), row.getString(9));
                assertEquals(rs1.getDate(11).getTime() / 1000, row.getDate(10).getTime() / 1000);
                assertEquals(rs1.getDate(12).getTime() / 1000, row.getDate(11).getTime() / 1000);
                assertEquals(rs1.getTimestamp(13).getTime() / 1000, row.getTimestamp(12).getTime() / 1000);

                // Row2
                assertTrue(iter.hasNext());
                assertTrue(rs1.next());
                // Row1
                row = iter.next();

                assertEquals(row.getCellType(0), CellType.INTEGER);
                assertEquals(row.getCellType(1), CellType.BLANK);
                assertEquals(row.getCellType(2), CellType.BLANK);
                assertEquals(row.getCellType(3), CellType.BLANK);
                assertEquals(row.getCellType(4), CellType.BLANK);
                assertEquals(row.getCellType(5), CellType.BLANK);
                assertEquals(row.getCellType(6), CellType.BLANK);
                assertEquals(row.getCellType(7), CellType.BLANK);
                assertEquals(row.getCellType(8), CellType.BLANK);
                assertEquals(row.getCellType(9), CellType.BLANK);
                assertEquals(row.getCellType(10), CellType.BLANK);
                assertEquals(row.getCellType(11), CellType.BLANK);
                assertEquals(row.getCellType(12), CellType.BLANK);

                assertEquals(rs1.getInt(1), (int) row.getInt(0));
            }
            rs1.close();
            ps1.close();
        }
    }

    @Test public void testSpecifyCoordinateWrite() throws SQLException, IOException {
        final String fileName = "test specify coordinate D4 ResultSheet.xlsx";
        try (
            Connection con = getConnection();
            PreparedStatement ps = con.prepareStatement("select * from types_test");
            ResultSet rs = ps.executeQuery()
        ) {
            new Workbook()
                .addSheet(new ResultSetSheet(rs).setStartCoordinate("D4"))
                .writeTo(defaultTestPath.resolve(fileName));
        }

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            Iterator<org.ttzero.excel.reader.Row> iter = reader.sheet(0).iterator();
            org.ttzero.excel.reader.Row firstRow = iter.next();
            assertNotNull(firstRow);
            assertEquals(firstRow.getRowNum(), 4);
            assertEquals(firstRow.getFirstColumnIndex(), 3);
        }
    }
}

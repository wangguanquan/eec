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
                assert iter.hasNext();
                org.ttzero.excel.reader.Row header = iter.next();
                assert "学号".equals(header.getString(0));
                assert "姓名".equals(header.getString(1));
                assert "年龄".equals(header.getString(2));
                assert "创建时间".equals(header.getString(3));
                assert "更新".equals(header.getString(4));
                while (rs.next()) {
                    assert iter.hasNext();
                    org.ttzero.excel.reader.Row row = iter.next();

                    assert rs.getInt(1) == row.getInt(0);
                    assert rs.getString(2).equals(row.getString(1));
                    assert rs.getInt(3) == row.getInt(2);
                    assert rs.getTimestamp(4) != null ? rs.getTimestamp(4).getTime() / 1000 == row.getTimestamp(3).getTime() / 1000 : row.getTimestamp(0) == null;
                    assert rs.getTimestamp(5) != null ? rs.getTimestamp(5).getTime() / 1000 == row.getTimestamp(4).getTime() / 1000 : row.getTimestamp(4) == null;
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
                            style = Styles.clearFill(style) | sst.addFill(new Fill(PatternType.solid, Color.yellow));
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
                assert iter.hasNext();
                org.ttzero.excel.reader.Row header = iter.next();
                assert "id".equals(header.getString(0));
                assert "name".equals(header.getString(1));
                assert "age".equals(header.getString(2));
                assert "create_date".equals(header.getString(3));
                assert "update_date".equals(header.getString(4));
                while (rs.next()) {
                    assert iter.hasNext();
                    org.ttzero.excel.reader.Row row = iter.next();

                    assert rs.getInt(1) == row.getInt(0);
                    assert rs.getString(2).equals(row.getString(1));
                    assert rs.getInt(3) == row.getInt(2);
                    assert rs.getTimestamp(4) != null ? rs.getTimestamp(4).getTime() / 1000 == row.getTimestamp(3).getTime() / 1000 : row.getTimestamp(0) == null;
                    assert rs.getTimestamp(5) != null ? rs.getTimestamp(5).getTime() / 1000 == row.getTimestamp(4).getTime() / 1000 : row.getTimestamp(4) == null;

                    Styles styles = row.getStyles();
                    int style = row.getCellStyle(2);
                    Fill fill = styles.getFill(style);
                    if (rs.getInt(3) > 14) {
                        assert fill != null && fill.getPatternType() == PatternType.solid && fill.getFgColor().equals(Color.yellow);
                    } else assert  fill == null || fill.getPatternType() == PatternType.none;
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
                assert iter.hasNext();
                org.ttzero.excel.reader.Row header = iter.next();
                assert "id".equals(header.getString(0));
                assert "name".equals(header.getString(1));
                assert "age".equals(header.getString(2));
                assert "create_date".equals(header.getString(3));
                assert "update_date".equals(header.getString(4));
                while (rs.next()) {
                    assert iter.hasNext();
                    org.ttzero.excel.reader.Row row = iter.next();

                    assert rs.getInt(1) == row.getInt(0);
                    assert rs.getString(2).equals(row.getString(1));
                    assert rs.getInt(3) == row.getInt(2);
                    assert rs.getTimestamp(4) != null ? rs.getTimestamp(4).getTime() / 1000 == row.getTimestamp(3).getTime() / 1000 : row.getTimestamp(0) == null;
                    assert rs.getTimestamp(5) != null ? rs.getTimestamp(5).getTime() / 1000 == row.getTimestamp(4).getTime() / 1000 : row.getTimestamp(4) == null;
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
                assert "Student".equals(sheet.getName());
                Iterator<Row> iter = sheet.iterator();
                assert iter.hasNext();
                org.ttzero.excel.reader.Row header = iter.next();
                assert "id".equals(header.getString(0));
                assert "name".equals(header.getString(1));
                assert "age".equals(header.getString(2));
                assert "create_date".equals(header.getString(3));
                assert "update_date".equals(header.getString(4));
                while (rs.next()) {
                    assert iter.hasNext();
                    org.ttzero.excel.reader.Row row = iter.next();

                    assert rs.getInt(1) == row.getInt(0);
                    assert rs.getString(2).equals(row.getString(1));
                    assert rs.getInt(3) == row.getInt(2);
                    assert rs.getTimestamp(4) != null ? rs.getTimestamp(4).getTime() / 1000 == row.getTimestamp(3).getTime() / 1000 : row.getTimestamp(0) == null;
                    assert rs.getTimestamp(5) != null ? rs.getTimestamp(5).getTime() / 1000 == row.getTimestamp(4).getTime() / 1000 : row.getTimestamp(4) == null;
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
                assert "Student".equals(sheet.getName());
                Iterator<Row> iter = sheet.iterator();
                assert iter.hasNext();
                org.ttzero.excel.reader.Row header = iter.next();
                assert "ID".equals(header.getString(0));
                assert "NAME".equals(header.getString(1));
                assert "AGE".equals(header.getString(2));
                assert "CREATE_DATE".equals(header.getString(3));
                assert "UPDATE_DATE".equals(header.getString(4));
                while (rs.next()) {
                    assert iter.hasNext();
                    org.ttzero.excel.reader.Row row = iter.next();

                    assert rs.getInt(1) == row.getInt(0);
                    assert rs.getString(2).equals(row.getString(1));
                    assert rs.getInt(3) == row.getInt(2);
                    assert rs.getTimestamp(4) != null ? rs.getTimestamp(4).getTime() / 1000 == row.getTimestamp(3).getTime() / 1000 : row.getTimestamp(0) == null;
                    assert rs.getTimestamp(5) != null ? rs.getTimestamp(5).getTime() / 1000 == row.getTimestamp(4).getTime() / 1000 : row.getTimestamp(4) == null;
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
                .addSheet(new ResultSetSheet("Student", WaterMark.of(author)
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
                assert "Student".equals(sheet.getName());
                Iterator<Row> iter = sheet.iterator();
                assert iter.hasNext();
                org.ttzero.excel.reader.Row header = iter.next();
                assert "ID".equals(header.getString(0));
                assert "NAME".equals(header.getString(1));
                assert "AGE".equals(header.getString(2));
                assert "CREATE_DATE".equals(header.getString(3));
                assert "UPDATE_DATE".equals(header.getString(4));
                while (rs.next()) {
                    assert iter.hasNext();
                    org.ttzero.excel.reader.Row row = iter.next();

                    assert rs.getInt(1) == row.getInt(0);
                    assert rs.getString(2).equals(row.getString(1));
                    assert rs.getInt(3) == row.getInt(2);
                    assert rs.getTimestamp(4) != null ? rs.getTimestamp(4).getTime() / 1000 == row.getTimestamp(3).getTime() / 1000 : row.getTimestamp(0) == null;
                    assert rs.getTimestamp(5) != null ? rs.getTimestamp(5).getTime() / 1000 == row.getTimestamp(4).getTime() / 1000 : row.getTimestamp(4) == null;
                }

                // Water Mark
                List<Drawings.Picture> pictures = sheet.listPictures();
                assert pictures.size() == 1;
                assert pictures.get(0).isBackground();
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
                assert iter.hasNext();
                org.ttzero.excel.reader.Row header = iter.next();
                assert "id".equals(header.getString(0));
                assert "name".equals(header.getString(1));
                assert "age".equals(header.getString(2));
                assert "create_date".equals(header.getString(3));
                assert "update_date".equals(header.getString(4));
                while (rs.next()) {
                    assert iter.hasNext();
                    org.ttzero.excel.reader.Row row = iter.next();

                    assert rs.getInt(1) == row.getInt(0);
                    assert rs.getString(2).equals(row.getString(1));
                    assert rs.getInt(3) == row.getInt(2);
                    assert rs.getTimestamp(4) != null ? rs.getTimestamp(4).getTime() / 1000 == row.getTimestamp(3).getTime() / 1000 : row.getTimestamp(0) == null;
                    assert rs.getTimestamp(5) != null ? rs.getTimestamp(5).getTime() / 1000 == row.getTimestamp(4).getTime() / 1000 : row.getTimestamp(4) == null;
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
                assert "Student".equals(sheet.getName());
                Iterator<Row> iter = sheet.iterator();
                assert iter.hasNext();
                org.ttzero.excel.reader.Row header = iter.next();
                assert "id".equals(header.getString(0));
                assert "name".equals(header.getString(1));
                assert "age".equals(header.getString(2));
                assert "create_date".equals(header.getString(3));
                assert "update_date".equals(header.getString(4));
                while (rs.next()) {
                    assert iter.hasNext();
                    org.ttzero.excel.reader.Row row = iter.next();

                    assert rs.getInt(1) == row.getInt(0);
                    assert rs.getString(2).equals(row.getString(1));
                    assert rs.getInt(3) == row.getInt(2);
                    assert rs.getTimestamp(4) != null ? rs.getTimestamp(4).getTime() / 1000 == row.getTimestamp(3).getTime() / 1000 : row.getTimestamp(0) == null;
                    assert rs.getTimestamp(5) != null ? rs.getTimestamp(5).getTime() / 1000 == row.getTimestamp(4).getTime() / 1000 : row.getTimestamp(4) == null;
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
                assert iter.hasNext();
                org.ttzero.excel.reader.Row header = iter.next();
                assert "ID".equals(header.getString(0));
                assert "NAME".equals(header.getString(1));
                assert "AGE".equals(header.getString(2));
                assert "CREATE_DATE".equals(header.getString(3));
                assert "UPDATE_DATE".equals(header.getString(4));
                while (rs.next()) {
                    assert iter.hasNext();
                    org.ttzero.excel.reader.Row row = iter.next();

                    assert rs.getInt(1) == row.getInt(0);
                    assert rs.getString(2).equals(row.getString(1));
                    assert rs.getInt(3) == row.getInt(2);
                    assert rs.getTimestamp(4) != null ? rs.getTimestamp(4).getTime() / 1000 == row.getTimestamp(3).getTime() / 1000 : row.getTimestamp(0) == null;
                    assert rs.getTimestamp(5) != null ? rs.getTimestamp(5).getTime() / 1000 == row.getTimestamp(4).getTime() / 1000 : row.getTimestamp(4) == null;
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
                assert "Student".equals(sheet.getName());
                Iterator<Row> iter = sheet.iterator();
                assert iter.hasNext();
                org.ttzero.excel.reader.Row header = iter.next();
                assert "ID".equals(header.getString(0));
                assert "NAME".equals(header.getString(1));
                assert "AGE".equals(header.getString(2));
                assert "CREATE_DATE".equals(header.getString(3));
                assert "UPDATE_DATE".equals(header.getString(4));
                while (rs.next()) {
                    assert iter.hasNext();
                    org.ttzero.excel.reader.Row row = iter.next();

                    assert rs.getInt(1) == row.getInt(0);
                    assert rs.getString(2).equals(row.getString(1));
                    assert rs.getInt(3) == row.getInt(2);
                    assert rs.getTimestamp(4) != null ? rs.getTimestamp(4).getTime() / 1000 == row.getTimestamp(3).getTime() / 1000 : row.getTimestamp(0) == null;
                    assert rs.getTimestamp(5) != null ? rs.getTimestamp(5).getTime() / 1000 == row.getTimestamp(4).getTime() / 1000 : row.getTimestamp(4) == null;
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
                .addSheet(new ResultSetSheet(rs, WaterMark.of(author)
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
                assert iter.hasNext();
                org.ttzero.excel.reader.Row header = iter.next();
                assert "ID".equals(header.getString(0));
                assert "NAME".equals(header.getString(1));
                assert "AGE".equals(header.getString(2));
                assert "CREATE_DATE".equals(header.getString(3));
                assert "UPDATE_DATE".equals(header.getString(4));
                while (rs.next()) {
                    assert iter.hasNext();
                    org.ttzero.excel.reader.Row row = iter.next();

                    assert rs.getInt(1) == row.getInt(0);
                    assert rs.getString(2).equals(row.getString(1));
                    assert rs.getInt(3) == row.getInt(2);
                    assert rs.getTimestamp(4) != null ? rs.getTimestamp(4).getTime() / 1000 == row.getTimestamp(3).getTime() / 1000 : row.getTimestamp(0) == null;
                    assert rs.getTimestamp(5) != null ? rs.getTimestamp(5).getTime() / 1000 == row.getTimestamp(4).getTime() / 1000 : row.getTimestamp(4) == null;
                }

                // Water Mark
                List<Drawings.Picture> pictures = sheet.listPictures();
                assert pictures.size() == 1;
                assert pictures.get(0).isBackground();
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
                .addSheet(new ResultSetSheet("Student", rs, WaterMark.of(author)
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
                assert iter.hasNext();
                org.ttzero.excel.reader.Row header = iter.next();
                assert "ID".equals(header.getString(0));
                assert "NAME".equals(header.getString(1));
                assert "AGE".equals(header.getString(2));
                assert "CREATE_DATE".equals(header.getString(3));
                assert "UPDATE_DATE".equals(header.getString(4));
                while (rs.next()) {
                    assert iter.hasNext();
                    org.ttzero.excel.reader.Row row = iter.next();

                    assert rs.getInt(1) == row.getInt(0);
                    assert rs.getString(2).equals(row.getString(1));
                    assert rs.getInt(3) == row.getInt(2);
                    assert rs.getTimestamp(4) != null ? rs.getTimestamp(4).getTime() / 1000 == row.getTimestamp(3).getTime() / 1000 : row.getTimestamp(0) == null;
                    assert rs.getTimestamp(5) != null ? rs.getTimestamp(5).getTime() / 1000 == row.getTimestamp(4).getTime() / 1000 : row.getTimestamp(4) == null;
                }

                // Water Mark
                List<Drawings.Picture> pictures = sheet.listPictures();
                assert pictures.size() == 1;
                assert pictures.get(0).isBackground();
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
                .addSheet(new ResultSetSheet("Student", rs, WaterMark.of(author)
                    , new Column("ID", String.class)  // Integer in database
                    , new Column("NAME", String.class)
                    , new Column("AGE", String.class) // // Integer in database
                    , new Column("CREATE_DATE", String.class)
                    , new Column("UPDATE_DATE", String.class)
                ))
                .writeTo(defaultTestPath.resolve(fileName));

            PreparedStatement ps1 = con.prepareStatement(sql);
            ResultSet rs1 = ps1.executeQuery();
            try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
                org.ttzero.excel.reader.Sheet sheet = reader.sheet(0);
                Iterator<Row> iter = sheet.iterator();
                assert iter.hasNext();
                org.ttzero.excel.reader.Row header = iter.next();
                assert "ID".equals(header.getString(0));
                assert "NAME".equals(header.getString(1));
                assert "AGE".equals(header.getString(2));
                assert "CREATE_DATE".equals(header.getString(3));
                assert "UPDATE_DATE".equals(header.getString(4));
                while (rs.next()) {
                    assert iter.hasNext();
                    org.ttzero.excel.reader.Row row = iter.next();

                    assert row.getCellType(0) == CellType.STRING;
                    assert row.getCellType(1) == CellType.STRING;
                    assert row.getCellType(2) == CellType.STRING;
                    assert row.getCellType(3) == CellType.STRING;
                    assert row.getCellType(4) == CellType.STRING || row.getCellType(4) == CellType.BLANK;

                    assert rs.getInt(1) == row.getInt(0);
                    assert rs.getString(2).equals(row.getString(1));
                    assert rs.getInt(3) == row.getInt(2);
                    assert rs.getTimestamp(4) != null ? rs.getTimestamp(4).getTime() / 1000 == row.getTimestamp(3).getTime() / 1000 : row.getTimestamp(0) == null;
                    assert rs.getTimestamp(5) != null ? rs.getTimestamp(5).getTime() / 1000 == row.getTimestamp(4).getTime() / 1000 : row.getTimestamp(4) == null;
                }

                // Water Mark
                List<Drawings.Picture> pictures = sheet.listPictures();
                assert pictures.size() == 1;
                assert pictures.get(0).isBackground();
            }
            rs1.close();
            ps1.close();
        }
    }


}

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

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertTrue;

/**
 * @author guanquan.wang at 2019-04-29 15:16
 */
public class ResultSetPagingTest extends SQLWorkbookTest {
    @Test public void testPaging() throws SQLException, IOException {
        String fileName = "result set pagingt.xlsx",
            sql = "select id, name, age, create_date, update_date from student";

        try (Connection con = getConnection()) {
            PreparedStatement ps = con.prepareStatement(sql);
            ResultSet rs = ps.executeQuery();
            Workbook workbook = new Workbook()
                .addSheet(new ResultSetSheet(
                    new Column("学号", int.class)
                    , new Column("姓名", String.class)
                    , new Column("年龄", int.class)
                    , new Column("创建时间", Timestamp.class)
                    , new Column("更新", Timestamp.class)
                ).setResultSet(rs))
            .setWorkbookWriter(new ReLimitXMLWorkbookWriter());
            workbook.writeTo(defaultTestPath.resolve(fileName));
            rs.close();
            ps.close();

            int rowLimit = workbook.getSheetAt(0).getSheetWriter().getRowLimit();

            ps = con.prepareStatement("select count(*) from student");
            rs = ps.executeQuery();
            int count = rs.getInt(1);
            rs.close();
            ps.close();

            ps = con.prepareStatement(sql);
            rs = ps.executeQuery();
            try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
                assertEquals(reader.getSheetCount(), (count % (rowLimit - 1) > 0 ? count / (rowLimit - 1) + 1 : count / (rowLimit - 1))); // Include header row

                for (int i = 0, len = reader.getSheetCount(); i < len; i++) {
                    Iterator<Row> iter = reader.sheet(i).iterator();
                    assertTrue(iter.hasNext());
                    org.ttzero.excel.reader.Row header = iter.next();
                    assertEquals("学号", header.getString(0));
                    assertEquals("姓名", header.getString(1));
                    assertEquals("年龄", header.getString(2));
                    assertEquals("创建时间", header.getString(3));
                    assertEquals("更新", header.getString(4));
                    int x = 1;
                    while (rs.next()) {
                        assertTrue(iter.hasNext());
                        org.ttzero.excel.reader.Row row = iter.next();

                        assertEquals(rs.getInt(1), (int) row.getInt(0));
                        assertEquals(rs.getString(2), row.getString(1));
                        assertEquals(rs.getInt(3), (int) row.getInt(2));
                        assertTrue(rs.getTimestamp(4) != null ? rs.getTimestamp(4).getTime() / 1000 == row.getTimestamp(3).getTime() / 1000 : row.getTimestamp(0) == null);
                        assertTrue(rs.getTimestamp(5) != null ? rs.getTimestamp(5).getTime() / 1000 == row.getTimestamp(4).getTime() / 1000 : row.getTimestamp(4) == null);

                        if (++x >= rowLimit) break;
                    }
                }
            }
            rs.close();
            ps.close();
        }
    }


    @Test public void testStyleDesignPaging() throws SQLException, IOException {
        String fileName = "test global style design for ResultSet Paging.xlsx",
            sql = "select id, name, age, create_date, update_date from student";

        try (Connection con = getConnection()) {
            PreparedStatement ps = con.prepareStatement(sql);
            ResultSet rs = ps.executeQuery();
            Workbook workbook = new Workbook()
                .addSheet(new ResultSetSheet().setResultSet(rs).setStyleProcessor((rst, style, sst)->{
                    try {
                        if (rst.getInt("age") > 14) {
                            style = sst.modifyFill(style, new Fill(PatternType.solid, Color.yellow));
                        }
                    } catch (SQLException ex) {
                        // Ignore
                    }
                    return style;
                }))
                .setWorkbookWriter(new ReLimitXMLWorkbookWriter());
            workbook.writeTo(defaultTestPath.resolve(fileName));
            rs.close();
            ps.close();

            int rowLimit = workbook.getSheetAt(0).getSheetWriter().getRowLimit();

            ps = con.prepareStatement("select count(*) from student");
            rs = ps.executeQuery();
            int count = rs.getInt(1);
            rs.close();
            ps.close();

            ps = con.prepareStatement(sql);
            rs = ps.executeQuery();
            try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
                assertEquals(reader.getSheetCount(), (count % (rowLimit - 1) > 0 ? count / (rowLimit - 1) + 1 : count / (rowLimit - 1))); // Include header row

                for (int i = 0, len = reader.getSheetCount(); i < len; i++) {
                    Iterator<Row> iter = reader.sheet(i).iterator();
                    assertTrue(iter.hasNext());
                    org.ttzero.excel.reader.Row header = iter.next();
                    assertEquals("id", header.getString(0));
                    assertEquals("name", header.getString(1));
                    assertEquals("age", header.getString(2));
                    assertEquals("create_date", header.getString(3));
                    assertEquals("update_date", header.getString(4));
                    int x = 1;
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
                        } else assertTrue(fill == null || fill.getPatternType() == PatternType.none);

                        if (++x >= rowLimit) break;
                    }
                }
            }
            rs.close();
            ps.close();
        }
    }

    @Test public void testSpecifyCoordinateWrite() throws SQLException, IOException {
        String fileName = "test specify coordinate D5 ResultSheet paging.xlsx",
            sql = "select id, name, age, create_date, update_date from student";

        try (Connection con = getConnection()) {
            PreparedStatement ps = con.prepareStatement(sql);
            ResultSet rs = ps.executeQuery();
            Workbook workbook = new Workbook()
                .addSheet(new ResultSetSheet().setResultSet(rs).setStyleProcessor((rst, style, sst) -> {
                    try {
                        if (rst.getInt("age") > 14) {
                            style = sst.modifyFill(style, new Fill(PatternType.solid, Color.yellow));
                        }
                    } catch (SQLException ex) {
                        // Ignore
                    }
                    return style;
                }).setStartCoordinate("D5"))
                .setWorkbookWriter(new ReLimitXMLWorkbookWriter());
            workbook.writeTo(defaultTestPath.resolve(fileName));
        }

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            reader.sheets().forEach(sheet -> {
                Iterator<org.ttzero.excel.reader.Row> iter = sheet.iterator();
                org.ttzero.excel.reader.Row firstRow = iter.next();
                assertNotNull(firstRow);
                assertEquals(firstRow.getRowNum(), 5);
                assertEquals(firstRow.getFirstColumnIndex(), 3);
            });
        }
    }
}

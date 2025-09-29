/*
 * Copyright (c) 2017-2022, guanquan.wang@yandex.com All Rights Reserved.
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
import org.ttzero.excel.entity.e7.XMLWorksheetWriter;
import org.ttzero.excel.reader.ExcelReader;
import org.ttzero.excel.reader.HeaderRow;
import org.ttzero.excel.reader.Row;
import org.ttzero.excel.reader.Sheet;

import java.io.IOException;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Timestamp;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;
import static org.ttzero.excel.entity.ListMapSheetTest.createTestData;

/**
 * @author guanquan.wang at 2022-08-02 19:17
 */
public class GridLinesTest extends SQLWorkbookTest {
    @Test public void testListSheet() throws IOException {
        String fileName = "ListSheet ignore grid lines.xlsx";
        List<ListObjectSheetTest.Item> expectList = ListObjectSheetTest.Item.randomTestData();
        new Workbook()
            .addSheet(new ListSheet<>(expectList).hideGridLines())
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            List<ListObjectSheetTest.Item> list = reader.sheet(0).bind(ListObjectSheetTest.Item.class, 1).rows().map(row -> (ListObjectSheetTest.Item) row.get()).collect(Collectors.toList());
            assertEquals(expectList.size(), list.size());
            for (int i = 0, len = expectList.size(); i < len; i++) {
                ListObjectSheetTest.Item expect = expectList.get(i), e = list.get(i);
                assertEquals(expect, e);
            }
        }
    }

    @Test public void testListSheetPaging() throws IOException {
        String fileName = "ListSheet Paging ignore grid lines.xlsx";
        List<ListObjectSheetTest.Item> expectList = ListObjectSheetTest.Item.randomTestData();
        Workbook workbook = new Workbook()
            .setAutoSize(true)
            .addSheet(new ListSheet<>(expectList).hideGridLines()
                .setSheetWriter(new XMLWorksheetWriter() {
                @Override
                public int getRowLimit() {
                    return 10;
                }
            }));
        workbook.writeTo(defaultTestPath.resolve(fileName));

        int count = expectList.size(), rowLimit = workbook.getSheetAt(0).getSheetWriter().getRowLimit();
        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            assertEquals(reader.getSheetCount(), (count % (rowLimit - 1) > 0 ? count / (rowLimit - 1) + 1 : count / (rowLimit - 1))); // Include header row

            for (int i = 0, len = reader.getSheetCount(), a = 0; i < len; i++) {
                Sheet sheet = reader.sheet(i).header(1).bind(ListObjectSheetTest.Item.class);
                org.ttzero.excel.reader.HeaderRow header = (HeaderRow) sheet.getHeader();
                assertEquals("id", header.get(0));
                assertEquals("name", header.get(1));
                Iterator<Row> iter = sheet.iterator();
                while (iter.hasNext()) {
                    ListObjectSheetTest.Item expect = expectList.get(a++), e = iter.next().get();
                    assertEquals(expect, e);
                }
            }
        }
    }

    @Test public void testListMapSheet() throws IOException {
        String fileName = "ListMapSheet ignore grid lines.xlsx";
        List<Map<String, Object>> expectList = createTestData();
        new Workbook()
            .addSheet(new ListMapSheet<>(expectList).hideGridLines())
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            List<Map<String, ?>> list = reader.sheet(0).dataRows().map(Row::toMap).collect(Collectors.toList());
            assertEquals(expectList.size(), list.size());
            for (int i = 0, len = expectList.size(); i < len; i++) {
                Map<String, ?> expect = expectList.get(i), e = list.get(i);
                assertEquals(expect, e);
            }
        }
    }

    @Test public void testStatementSheet() throws SQLException, IOException {
        String fileName = "StatementSheet ignore grid lines.xlsx";
        String sql = "select id, name, age, create_date, update_date from student order by age";
        try (Connection con = getConnection()) {
            new Workbook()
                .addSheet(new StatementSheet(con, sql
                    , new Column("学号", int.class)
                    , new Column("姓名", String.class)
                    , new Column("年龄", int.class)
                    , new Column("创建时间", Timestamp.class).setColIndex(0)
                    , new Column("更新", Timestamp.class)
                ).hideGridLines()).writeTo(defaultTestPath.resolve(fileName));

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

                    assertTrue(rs.getTimestamp(4) != null ? rs.getTimestamp(4).getTime() / 1000 == row.getTimestamp(0).getTime() / 1000 : row.getTimestamp(0) == null);
                    assertEquals(rs.getInt(1), (long) row.getInt(1));
                    assertEquals(rs.getString(2), row.getString(2));
                    assertEquals(rs.getInt(3), (long) row.getInt(3));
                    assertTrue(rs.getTimestamp(5) != null ? rs.getTimestamp(5).getTime() / 1000 == row.getTimestamp(4).getTime() / 1000 : row.getTimestamp(4) == null);
                }
            }
            rs.close();
            ps.close();
        }
    }

    @Test public void testResultSetSheet() throws SQLException, IOException {
        String fileName = "ResultSetSheet ignore grid lines.xlsx";
        String sql = "select id, name, age, create_date, update_date from student limit 10";
        try (
            Connection con = getConnection();
            PreparedStatement ps = con.prepareStatement(sql);
            ResultSet rs = ps.executeQuery()
        ) {
            new Workbook()
                .addSheet(new ResultSetSheet(rs
                    , new Column("学号", int.class)
                    , new Column("姓名", String.class)
                    , new Column("年龄", Integer.class)
                    , new Column("创建时间", Timestamp.class)
                    , new Column("更新", Timestamp.class)
                ).hideGridLines()).writeTo(defaultTestPath.resolve(fileName));

            PreparedStatement ps1 = con.prepareStatement(sql);
            ResultSet rs1 = ps.executeQuery();
            try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
                Iterator<org.ttzero.excel.reader.Row> iter = reader.sheet(0).iterator();
                assertTrue(iter.hasNext());
                org.ttzero.excel.reader.Row header = iter.next();
                assertEquals("学号", header.getString(0));
                assertEquals("姓名", header.getString(1));
                assertEquals("年龄", header.getString(2));
                assertEquals("创建时间", header.getString(3));
                assertEquals("更新", header.getString(4));

                while (rs1.next()) {
                    assertTrue(iter.hasNext());
                    org.ttzero.excel.reader.Row row = iter.next();

                    assertEquals(rs1.getInt(1), (long) row.getInt(0));
                    assertEquals(rs1.getString(2), row.getString(1));
                    assertEquals(rs1.getInt(3), (long) row.getInt(2));
                    assertTrue(rs1.getTimestamp(4) != null ? rs.getTimestamp(4).getTime() / 1000 == row.getTimestamp(3).getTime() / 1000 : row.getTimestamp(3) == null);
                    assertTrue(rs1.getTimestamp(5) != null ? rs.getTimestamp(5).getTime() / 1000 == row.getTimestamp(4).getTime() / 1000 : row.getTimestamp(4) == null);
                }
            }
            rs1.close();
            ps1.close();
        }
    }
}

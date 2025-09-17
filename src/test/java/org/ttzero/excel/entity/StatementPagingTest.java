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
import org.ttzero.excel.reader.ExcelReader;
import org.ttzero.excel.reader.Row;

import java.io.IOException;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.Iterator;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;

/**
 * @author guanquan.wang at 2019-04-28 22:47
 */
public class StatementPagingTest extends SQLWorkbookTest {
    @Test public void testPaging() throws SQLException, IOException {
        try (Connection con = getConnection()) {
            String fileName = "statement paging.xlsx",
                sql = "select id, name, age, create_date, update_date from student";

            Workbook workbook = new Workbook()
                .addSheet(new StatementSheet(con, sql))
                .setWorkbookWriter(new ReLimitXMLWorkbookWriter());
            workbook.writeTo(defaultTestPath.resolve(fileName));

            int rowLimit = workbook.getSheet(0).getSheetWriter().getRowLimit();

            PreparedStatement ps = con.prepareStatement("select count(*) from student");
            ResultSet rs = ps.executeQuery();
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

                        if (++x >= rowLimit) break;
                    }
                }
            }
            rs.close();
            ps.close();
        }
    }
}

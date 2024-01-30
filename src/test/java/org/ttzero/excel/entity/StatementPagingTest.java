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

            int rowLimit = workbook.getSheetAt(0).getSheetWriter().getRowLimit();

            PreparedStatement ps = con.prepareStatement("select count(*) from student");
            ResultSet rs = ps.executeQuery();
            int count = rs.getInt(1);
            rs.close();
            ps.close();

            ps = con.prepareStatement(sql);
            rs = ps.executeQuery();
            try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
                assert reader.getSheetCount() == (count % (rowLimit - 1) > 0 ? count / (rowLimit - 1) + 1 : count / (rowLimit - 1)); // Include header row

                for (int i = 0, len = reader.getSheetCount(); i < len; i++) {
                    Iterator<Row> iter = reader.sheet(i).iterator();
                    assert iter.hasNext();
                    org.ttzero.excel.reader.Row header = iter.next();
                    assert "id".equals(header.getString(0));
                    assert "name".equals(header.getString(1));
                    assert "age".equals(header.getString(2));
                    assert "create_date".equals(header.getString(3));
                    assert "update_date".equals(header.getString(4));
                    int x = 1;
                    while (rs.next()) {
                        assert iter.hasNext();
                        org.ttzero.excel.reader.Row row = iter.next();

                        assert rs.getInt(1) == row.getInt(0);
                        assert rs.getString(2).equals(row.getString(1));
                        assert rs.getInt(3) == row.getInt(2);
                        assert rs.getTimestamp(4) != null ? rs.getTimestamp(4).getTime() / 1000 == row.getTimestamp(3).getTime() / 1000 : row.getTimestamp(0) == null;
                        assert rs.getTimestamp(5) != null ? rs.getTimestamp(5).getTime() / 1000 == row.getTimestamp(4).getTime() / 1000 : row.getTimestamp(4) == null;

                        if (++x >= rowLimit) break;
                    }
                }
            }
            rs.close();
            ps.close();
        }
    }
}

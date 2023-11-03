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
                assert reader.getSize() == (count % (rowLimit - 1) > 0 ? count / (rowLimit - 1) + 1 : count / (rowLimit - 1)); // Include header row

                for (int i = 0, len = reader.getSize(); i < len; i++) {
                    Iterator<Row> iter = reader.sheet(i).iterator();
                    assert iter.hasNext();
                    org.ttzero.excel.reader.Row header = iter.next();
                    assert "学号".equals(header.getString(0));
                    assert "姓名".equals(header.getString(1));
                    assert "年龄".equals(header.getString(2));
                    assert "创建时间".equals(header.getString(3));
                    assert "更新".equals(header.getString(4));
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
                assert reader.getSize() == (count % (rowLimit - 1) > 0 ? count / (rowLimit - 1) + 1 : count / (rowLimit - 1)); // Include header row

                for (int i = 0, len = reader.getSize(); i < len; i++) {
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

                        Styles styles = row.getStyles();
                        int style = row.getCellStyle(2);
                        Fill fill = styles.getFill(style);
                        if (rs.getInt(3) > 14) {
                            assert fill != null && fill.getPatternType() == PatternType.solid && fill.getFgColor().equals(Color.yellow);
                        } else assert  fill == null || fill.getPatternType() == PatternType.none;

                        if (++x >= rowLimit) break;
                    }
                }
            }
            rs.close();
            ps.close();
        }
    }
}

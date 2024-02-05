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
import org.ttzero.excel.reader.Sheet;

import java.awt.Color;
import java.io.IOException;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;

/**
 * @author guanquan.wang at 2019-05-01 19:34
 */
public class MultiWorksheetTest extends SQLWorkbookTest {

    @Test public void testMultiWorksheet() throws IOException {
        List<Map<String, ?>> sheet1Data = ListMapSheetTest.createTestData(), sheet2Data = ListMapSheetTest.createAllTypeData();

        new Workbook()
                .setAutoSize(true)
                // The first worksheet
                .addSheet(new ListMapSheet("E", sheet1Data))
                // The other worksheet
                .addSheet(new ListMapSheet("All type", sheet2Data))
                .writeTo(defaultTestPath.resolve("test multi worksheet.xlsx"));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("test multi worksheet.xlsx"))) {
            Sheet sheet0 = reader.sheet(0);
            assertEquals("E", sheet0.getName());
            Iterator<org.ttzero.excel.reader.Row> iter = sheet0.iterator();
            // Check header
            assertTrue(iter.hasNext());
            org.ttzero.excel.reader.Row header = iter.next();
            assertEquals("id", header.getString(0));
            assertEquals("name", header.getString(1));
            for (Map<String, ?> expect : sheet1Data) {
                assertTrue(iter.hasNext());
                org.ttzero.excel.reader.Row row = iter.next();

                assertEquals((int) row.getInt(0), Integer.parseInt(expect.get("id").toString()));
                assertEquals(row.getString(1), expect.get("name"));
            }

            Sheet sheet1 = reader.sheet(1);
            assertEquals("All type", sheet1.getName());
            List<Map<String, Object>> list2 = sheet1.dataRows().map(Row::toMap).collect(Collectors.toList());
            assertEquals(list2.size(), sheet2Data.size());

            assertEquals(String.join(",", sheet2Data.get(0).keySet()), String.join(",", list2.get(0).keySet()));

//            for (int i = 0, len = list2.size(); i < len; i++) {
//                Map<String, ?> expectMap = sheet2Data.get(i), map = list2.get(i);
//                assertEquals(expectMap, map);
//            }
        }
    }

    @Test public void testMultiDataSource() throws SQLException, IOException {
        List<Map<String, ?>> sheet1Data =  ListMapSheetTest.createAllTypeData();
        List<ListObjectSheetTest.Item> sheet2Data = ListObjectSheetTest.Item.randomTestData();
        String sql3 = "select id, name, age from student", sql5 = "select id, name, age from student order by age limit 10";
        List<ListObjectSheetTest.Student> sheet5Data = new ArrayList<>();
        try (
            Connection con = getConnection();
            PreparedStatement ps = con.prepareStatement(sql5);
            ResultSet rs = ps.executeQuery()
        ) {

            new Workbook()
                .setAutoSize(true)
                // List<Map>
                .addSheet(new ListMapSheet("ListMap", sheet1Data))
                // List<Object>
                .addSheet(new ListSheet<>("ListObject", sheet2Data))
                // Statement
                .addSheet(new StatementSheet("Statement", con, sql3
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
                // Empty
                .addSheet(new EmptySheet("Empty"))
                // ResultSet
                .addSheet(new ResultSetSheet("ResultSet", rs
                    , new Column("学号", int.class)
                    , new Column("姓名", String.class)
                    , new Column("年龄", int.class)
                ))
                // Customize
                .addSheet(new CustomizeDataSourceSheet("Customize") {
                    @Override
                    public List<ListObjectSheetTest.Student> more() {
                        List<ListObjectSheetTest.Student> sub = super.more();
                        if (sub != null) sheet5Data.addAll(sub);
                        return sub;
                    }
                })
                .writeTo(defaultTestPath.resolve("test multi dataSource worksheet.xlsx"));

            try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("test multi dataSource worksheet.xlsx"))) {
                assertEquals(reader.getSheetCount(), 6);

                // Sheet 0
                Sheet sheet0 = reader.sheet(0);
                assertEquals("ListMap", sheet0.getName());
                List<Map<String, Object>> list0 = sheet0.dataRows().map(Row::toMap).collect(Collectors.toList());
                assertEquals(list0.size(), sheet1Data.size());

                assertEquals(String.join(",", sheet1Data.get(0).keySet()), String.join(",", list0.get(0).keySet()));

//                for (int i = 0, len = list0.size(); i < len; i++) {
//                    Map<String, ?> expectMap = sheet1Data.get(i), map = list0.get(i);
//                    assertEquals(expectMap, map);
//                }

                // Sheet 1
                Sheet sheet1 = reader.sheet(1);
                assertEquals("ListObject", sheet1.getName());
                List<ListObjectSheetTest.Item> list1 = sheet1.dataRows().map(row -> row.to(ListObjectSheetTest.Item.class)).collect(Collectors.toList());
                assertEquals(list1.size(), sheet2Data.size());

                for (int i = 0, len = list1.size(); i < len; i++) {
                    ListObjectSheetTest.Item expect = sheet2Data.get(i), o = list1.get(i);
                    assertEquals(expect, o);
                }

                // Sheet 2
                Sheet sheet2 = reader.sheet(2);
                assertEquals("Statement", sheet2.getName());
                Iterator<org.ttzero.excel.reader.Row> iter = sheet2.iterator();
                assertTrue(iter.hasNext());
                // Header row
                org.ttzero.excel.reader.Row header = iter.next();
                assertEquals("学号", header.getString(0));
                assertEquals("姓名", header.getString(1));
                assertEquals("年龄", header.getString(2));

                PreparedStatement ps3 = con.prepareStatement(sql3);
                ResultSet rs3 = ps3.executeQuery();
                // Body rows
                while (rs3.next()) {
                    assertTrue(iter.hasNext());
                    org.ttzero.excel.reader.Row row = iter.next();

                    assertEquals(rs3.getInt(1), (int) row.getInt(0));
                    assertEquals(rs3.getString(2), row.getString(1));

                    Styles styles = row.getStyles();
                    int style = row.getCellStyle(2);
                    Fill fill = styles.getFill(style);

                    int age = rs3.getInt(3);
                    if (age > 14) {
                        assertEquals("高龄", row.getString(2));
                        assertTrue(fill != null && fill.getPatternType() == PatternType.solid && fill.getFgColor().equals(Color.orange));
                    } else {
                        assertEquals(age, (int) row.getInt(2));
                        assertTrue(fill == null || fill.getPatternType() == PatternType.none);
                    }
                }

                rs3.close();
                ps3.close();


                // Sheet 3
                Sheet sheet3 = reader.sheet(3);
                assertEquals("Empty", sheet3.getName());
                assertEquals(sheet3.rows().count(), 0L);

                // Sheet 4
                Sheet sheet4 = reader.sheet(4);
                assertEquals("ResultSet", sheet4.getName());
                Iterator<org.ttzero.excel.reader.Row> iter4 = sheet4.iterator();
                assertTrue(iter4.hasNext());
                // Header row
                org.ttzero.excel.reader.Row header4 = iter4.next();
                assertEquals("学号", header4.getString(0));
                assertEquals("姓名", header4.getString(1));
                assertEquals("年龄", header4.getString(2));

                PreparedStatement ps4 = con.prepareStatement(sql5);
                ResultSet rs4 = ps4.executeQuery();
                // Body rows
                while (rs4.next()) {
                    assertTrue(iter4.hasNext());
                    org.ttzero.excel.reader.Row row = iter4.next();

                    assertEquals(rs4.getInt(1), (int) row.getInt(0));
                    assertEquals(rs4.getString(2), row.getString(1));
                    assertEquals(rs4.getInt(3), (int) row.getInt(2));
                }

                rs4.close();
                ps4.close();

                // Sheet 5
                Sheet sheet5 = reader.sheet(5);
                assertEquals("Customize", sheet5.getName());
                List<ListObjectSheetTest.Student> sheet5ReadList = sheet5.dataRows().map(row -> row.to(ListObjectSheetTest.Student.class)).collect(Collectors.toList());
                assertEquals(sheet5Data.size(), sheet5ReadList.size());
                for (int i = 0, len = sheet5Data.size(); i < len; i++) {
                    ListObjectSheetTest.Student expect = sheet5Data.get(i), e = sheet5ReadList.get(i);
                    assertEquals(expect.getName(), e.getName());
                    assertEquals(expect.getScore(), e.getScore());
                }
            }
        }
    }
}

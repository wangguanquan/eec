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
import org.ttzero.excel.annotation.ExcelColumn;
import org.ttzero.excel.annotation.HeaderComment;
import org.ttzero.excel.entity.e7.XMLWorksheetWriter;
import org.ttzero.excel.entity.style.Styles;
import org.ttzero.excel.reader.Cell;
import org.ttzero.excel.reader.ExcelReader;
import org.ttzero.excel.entity.style.Font;
import org.ttzero.excel.entity.style.Horizontals;
import org.ttzero.excel.reader.HeaderRow;
import org.ttzero.excel.reader.Row;
import org.ttzero.excel.entity.style.NumFmt;
import org.ttzero.excel.reader.Sheet;

import java.awt.Color;
import java.io.IOException;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Time;
import java.sql.Timestamp;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.stream.Collectors;


/**
 * @author guanquan.wang at 2022-06-27 23:24
 */
public class MultiHeaderColumnsTest extends SQLWorkbookTest {
    @Test public void testRepeatAnnotations() throws IOException {
        List<RepeatableEntry> list = RepeatableEntry.randomTestData();
        new Workbook().setWaterMark(WaterMark.of("勿外传"))
            .setAutoSize(true)
            .addSheet(new ListSheet<>(list))
            .writeTo(defaultTestPath.resolve("Repeat Columns Annotation.xlsx"));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("Repeat Columns Annotation.xlsx"))) {
            List<RepeatableEntry> readList = reader.sheet(0).header(1, 4).bind(RepeatableEntry.class).rows()
                .map(row -> (RepeatableEntry) row.get()).collect(Collectors.toList());

            assert list.size() == readList.size();
            for (int i = 0, len = list.size(); i < len; i++)
                assert list.get(i).equals(readList.get(i));


            // Row to Map
            List<Map<String, Object>> mapList = reader.sheet(0).header(1, 4).rows().map(Row::toMap).collect(Collectors.toList());
            assert list.size() == mapList.size();
            for (int i = 0, len = list.size(); i < len; i++) {
                Map<String, Object> sub = mapList.get(i);
                RepeatableEntry src = list.get(i);

                assert sub.get("TOP:K:订单号").equals(src.orderNo);
                assert sub.get("TOP:K:A:收件人").equals(src.recipient);
                assert sub.get("TOP:收件地址:A:省").equals(src.province);
                assert sub.get("TOP:收件地址:A:市").equals(src.city);
                assert sub.get("TOP:收件地址:B:区").equals(src.area);
                assert sub.get("TOP:收件地址:B:详细地址").equals(src.detail);
            }
        }
    }

    @Test public void testPagingRepeatAnnotations() throws IOException {
        List<RepeatableEntry> expectList = RepeatableEntry.randomTestData(10000);
        IWorksheetWriter worksheetWriter;
        new Workbook().setAutoSize(true)
            .addSheet(new ListSheet<>(expectList).setSheetWriter(worksheetWriter = new XMLWorksheetWriter() {
                @Override
                public int getRowLimit() {
                    return 500;
                }
            })).writeTo(defaultTestPath.resolve("Repeat Paging Columns Annotation.xlsx"));

        int count = expectList.size(), rowLimit = worksheetWriter.getRowLimit();
        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("Repeat Paging Columns Annotation.xlsx"))) {
            assert reader.getSize() == (count % rowLimit > 0 ? count / rowLimit + 1 : count / (rowLimit - 1)); // Include header row

            for (int i = 0, len = reader.getSize(), a = 0; i < len; i++) {
                List<RepeatableEntry> list = reader.sheet(i).header(1, 4).bind(RepeatableEntry.class).rows().map(row -> (RepeatableEntry) row.get()).collect(Collectors.toList());
                if (i < len - 1) assert list.size() == rowLimit;
                else assert expectList.size() - rowLimit * (len - 1) == list.size();
                for (int j = 0; j < list.size(); j++) {
                    RepeatableEntry expect = expectList.get(a++), o = list.get(j);
                    assert expect.equals(o);
                }
            }
        }
    }

    @Test public void testMultiOrderColumnSpecifyOnColumn() throws IOException {
        List<ListObjectSheetTest.Student> expectList = ListObjectSheetTest.Student.randomTestData();
        new Workbook().setAutoSize(true)
            .addSheet(new ListSheet<>("期末成绩", expectList
                , new Column("共用表头").addSubColumn(new Column("学号", "id"))
                , new Column("共用表头").addSubColumn(new Column("姓名", "name"))
                , new Column("成绩", "score") {
                @Override
                public int getHeaderStyleIndex() {
                    return styles.of(styles.addFont(this.getFont()) | Horizontals.CENTER);
                }

                public Font getFont() {
                    return new Font("宋体", 12, Color.RED).bold();
                }
            }
            )).writeTo(defaultTestPath.resolve("Multi specify columns 2.xlsx"));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("Multi specify columns 2.xlsx"))) {
            Sheet sheet = reader.sheet(0);
            assert "期末成绩".equals(sheet.getName());
            List<Map<String, Object>> list = sheet.header(1, 2).rows().map(Row::toMap).collect(Collectors.toList());
            assert expectList.size() == list.size();
            for (int i = 0, len = expectList.size(); i < len; i++) {
                ListObjectSheetTest.Student expect = expectList.get(i);
                Map<String, Object> o = list.get(i);
                assert expect.getId() == Integer.parseInt(o.get("共用表头:学号").toString());
                assert expect.getName().equals(o.get("共用表头:姓名").toString());
                assert expect.getScore() == Integer.parseInt(o.get("成绩").toString());
            }

            Iterator<Row> iterator =  sheet.reset().iterator();
            Row firstRow = iterator.next();
            Styles styles = firstRow.getStyles();
            int style = firstRow.getCellStyle(2);
            Font font = styles.getFont(style);

            assert font.isBold();
            assert "宋体".equals(font.getName());
            assert Color.RED.equals(font.getColor());
            assert 12 == font.getSize();
        }
    }

    @Test public void testMultiOrderColumnSpecifyOnColumn3() throws IOException {
        List<ListObjectSheetTest.Student> expectList = ListObjectSheetTest.Student.randomTestData();
        new Workbook().setAutoSize(true)
            .addSheet(new ListSheet<>("期末成绩", expectList
                , new Column().addSubColumn(new ListSheet.EntryColumn("共用表头")).addSubColumn(new Column("学号", "id").setHeaderComment(new Comment("abc", "content")))
                , new ListSheet.EntryColumn("共用表头").addSubColumn(new Column("姓名", "name"))
                , new Column("成绩", "score")
            )).writeTo(defaultTestPath.resolve("Multi specify columns 3.xlsx"));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("Multi specify columns 3.xlsx"))) {
            Sheet sheet = reader.sheet(0);
            assert "期末成绩".equals(sheet.getName());
            List<Map<String, Object>> list = sheet.header(1, 3).rows().map(Row::toMap).collect(Collectors.toList());
            assert expectList.size() == list.size();
            for (int i = 0, len = expectList.size(); i < len; i++) {
                ListObjectSheetTest.Student expect = expectList.get(i);
                Map<String, Object> o = list.get(i);
                assert expect.getId() == Integer.parseInt(o.get("共用表头:学号").toString());
                assert expect.getName().equals(o.get("共用表头:姓名").toString());
                assert expect.getScore() == Integer.parseInt(o.get("成绩").toString());
            }
        }
    }

    @Test public void testResultSet() throws SQLException, IOException {
        String fileName = "Multi ResultSet columns 2.xlsx",
            sql = "select id, name, age, create_date, update_date from student order by age";

        try (Connection con = getConnection()) {
            new Workbook().setAutoSize(true)
                .addSheet(new StatementSheet(con, sql
                    , new Column("通用").setHeaderStyle(532550).addSubColumn(new Column("学号", int.class))
                    , new Column("通用").addSubColumn(new Column("姓名", String.class))
                    , new Column("通用").addSubColumn(new Column("年龄", int.class).setHeaderStyle(532550))
                    , new Column("创建时间", Timestamp.class)
                    , new Column("更新时间", Timestamp.class).setColIndex(1) // col 1
                ))
                .writeTo(defaultTestPath.resolve(fileName));

            PreparedStatement ps = con.prepareStatement(sql);
            ResultSet rs = ps.executeQuery();
            try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
                org.ttzero.excel.reader.Sheet sheet = reader.sheet(0);

                // Header row
                String[] headerNames = ((HeaderRow) sheet.header(1, 2).getHeader()).getNames();
                assert "通用:学号".equals(headerNames[0]);
                assert "更新时间".equals(headerNames[1]);
                assert "通用:姓名".equals(headerNames[2]);
                assert "通用:年龄".equals(headerNames[3]);
                assert "创建时间".equals(headerNames[4]);

                Iterator<org.ttzero.excel.reader.Row> iter = sheet.rows().iterator();
                // Body rows
                while (rs.next()) {
                    assert iter.hasNext();
                    org.ttzero.excel.reader.Row row = iter.next();

                    assert rs.getInt(1) == row.getInt(0);
                    assert rs.getTimestamp(5) != null ? rs.getTimestamp(5).getTime() / 1000 == row.getTimestamp(1).getTime() / 1000 : row.getTimestamp(1) == null;
                    assert rs.getString(2).equals(row.getString(2));
                    assert rs.getInt(3) == row.getInt(3);
                    assert rs.getTimestamp(4) != null ? rs.getTimestamp(4).getTime() / 1000 == row.getTimestamp(4).getTime() / 1000 : row.getTimestamp(4) == null;
                }
            }
            rs.close();
            ps.close();
        }
    }

    @Test public void testMultiHeaderAndSpecifyColIndex() throws SQLException, IOException {
        String fileName = "Multi Header And Specify Col-index.xlsx",
            sql = "select id, name, age, create_date, update_date from student limit 10";
        try (Connection con = getConnection()) {
            new Workbook().setAutoSize(true)
                .addSheet(new StatementSheet(con, sql
                    , new Column("通用").addSubColumn(new Column("学号", int.class))
                    , new Column("通用").addSubColumn(new Column("姓名", String.class).setColIndex(13))
                    , new Column("通用").addSubColumn(new Column("年龄", int.class).setColIndex(14))
                    , new Column("创建时间", Timestamp.class).setColIndex(15)
                    , new Column("更新时间", Timestamp.class).setColIndex(16)
                ))
                .writeTo(defaultTestPath.resolve(fileName));

            PreparedStatement ps = con.prepareStatement(sql);
            ResultSet rs = ps.executeQuery();
            try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
                org.ttzero.excel.reader.Sheet sheet = reader.sheet(0);

                // Header row
                String[] headerNames = ((HeaderRow) sheet.header(1, 2).getHeader()).getNames();
                assert "通用:学号".equals(headerNames[0]);
                assert "通用:姓名".equals(headerNames[13]);
                assert "通用:年龄".equals(headerNames[14]);
                assert "创建时间".equals(headerNames[15]);
                assert "更新时间".equals(headerNames[16]);

                Iterator<org.ttzero.excel.reader.Row> iter = sheet.rows().iterator();
                // Body rows
                while (rs.next()) {
                    assert iter.hasNext();
                    org.ttzero.excel.reader.Row row = iter.next();

                    assert rs.getInt(1) == row.getInt(0);
                    assert rs.getString(2).equals(row.getString(13));
                    assert rs.getInt(3) == row.getInt(14);
                    assert rs.getTimestamp(4) != null ? rs.getTimestamp(4).getTime() / 1000 == row.getTimestamp(15).getTime() / 1000 : row.getTimestamp(15) == null;
                    assert rs.getTimestamp(5) != null ? rs.getTimestamp(5).getTime() / 1000 == row.getTimestamp(16).getTime() / 1000 : row.getTimestamp(16) == null;
                }
            }
            rs.close();
            ps.close();
        }
    }

    @Test public void testRepeatAnnotations2() throws IOException {
        List<RepeatableEntry> list = RepeatableEntry.randomTestData();
        new Workbook()
            .addSheet(new ListSheet<>(list))
            .writeTo(defaultTestPath.resolve("Repeat Columns Annotation2.xlsx"));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("Repeat Columns Annotation2.xlsx"))) {
            List<RepeatableEntry> readList = reader.sheet(0).header(2, 4).bind(RepeatableEntry.class).rows()
                .map(row -> (RepeatableEntry) row.get()).collect(Collectors.toList());

            assert list.size() == readList.size();
            for (int i = 0, len = list.size(); i < len; i++)
                assert list.get(i).equals(readList.get(i));

            // Specify single header row
            org.ttzero.excel.reader.Row headerRow = new org.ttzero.excel.reader.Row() {};
            Cell[] cells = new Cell[6];
            cells[0] = new Cell((short) 1).setSv("订单号");
            cells[1] = new Cell((short) 2).setSv("收件人");
            cells[2] = new Cell((short) 3).setSv("省");
            cells[3] = new Cell((short) 4).setSv("市");
            cells[4] = new Cell((short) 5).setSv("区");
            cells[5] = new Cell((short) 6).setSv("详细地址");
            headerRow.setCells(cells);
            readList = reader.sheet(0).reset().header(4).bind(RepeatableEntry.class, new HeaderRow().with(headerRow))
                .rows().map(row -> (RepeatableEntry) row.get()).collect(Collectors.toList());
            assert list.size() == readList.size();
            for (int i = 0, len = list.size(); i < len; i++)
                assert list.get(i).equals(readList.get(i));

            // Specify 2 header rows
            org.ttzero.excel.reader.Row headerRow2 = new org.ttzero.excel.reader.Row() {};
            Cell[] cells2 = new Cell[6];
            cells2[0] = new Cell((short) 1).setSv("订单号");
            cells2[1] = new Cell((short) 2).setSv("A:收件人");
            cells2[2] = new Cell((short) 3).setSv("A:省");
            cells2[3] = new Cell((short) 4).setSv("A:市");
            cells2[4] = new Cell((short) 5).setSv("B:区");
            cells2[5] = new Cell((short) 6).setSv("B:详细地址");
            headerRow2.setCells(cells2);
            readList = reader.sheet(0).reset().header(4).bind(RepeatableEntry.class, new HeaderRow().with(2, headerRow2))
                .rows().map(row -> (RepeatableEntry) row.get()).collect(Collectors.toList());
            assert list.size() == readList.size();
            for (int i = 0, len = list.size(); i < len; i++)
                assert list.get(i).equals(readList.get(i));
        }
    }

    @Test public void testRepeatAnnotations3() throws IOException {
        List<RepeatableEntry3> list = RepeatableEntry3.randomTestData();
        new Workbook()
            .addSheet(new ListSheet<>(list))
            .writeTo(defaultTestPath.resolve("Repeat Columns Annotation3.xlsx"));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("Repeat Columns Annotation3.xlsx"))) {
            List<RepeatableEntry3> readList;

            // header row 4
//            readList = reader.sheet(0).header(4).bind(RepeatableEntry3.class).rows()
//                .map(row -> (RepeatableEntry3) row.get()).collect(Collectors.toList());
//
//            assert list.size() == readList.size();
//            for (int i = 0, len = list.size(); i < len; i++)
//                assert list.get(i).equals(readList.get(i));
//
//            // header rows 3-4
//            readList = reader.sheet(0).reset().header(3, 4).bind(RepeatableEntry3.class).rows()
//                .map(row -> (RepeatableEntry3) row.get()).collect(Collectors.toList());
//
//            assert list.size() == readList.size();
//            for (int i = 0, len = list.size(); i < len; i++)
//                assert list.get(i).equals(readList.get(i));
//
//            // header rows 2-4
//            readList = reader.sheet(0).reset().header(2, 4).bind(RepeatableEntry3.class).rows()
//                .map(row -> (RepeatableEntry3) row.get()).collect(Collectors.toList());
//
//            assert list.size() == readList.size();
//            for (int i = 0, len = list.size(); i < len; i++)
//                assert list.get(i).equals(readList.get(i));

            // header rows 1-4
            readList = reader.sheet(0).reset().header(1, 4).bind(RepeatableEntry3.class).rows()
                .map(row -> (RepeatableEntry3) row.get()).collect(Collectors.toList());

            assert list.size() == readList.size();
            for (int i = 0, len = list.size(); i < len; i++)
                assert list.get(i).equals(readList.get(i));
        }
    }

    @Test public void testAutoSizeAndHideCol() throws IOException {
        List<ListObjectSheetTest.Student> expectList = ListObjectSheetTest.Student.randomTestData();
        new Workbook().setAutoSize(true)
            .addSheet(new ListSheet<>("期末成绩", expectList
                , new Column().addSubColumn(new ListSheet.EntryColumn("共用表头")).addSubColumn(new Column("学号", "id").setHeaderComment(new Comment("abc", "content")))
                , new ListSheet.EntryColumn("共用表头").addSubColumn(new Column("姓名", "name").setColIndex(1000).hide())
                , new Column("成绩", "score")
            )).writeTo(defaultTestPath.resolve("Auto Size And Hide Column.xlsx"));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("Auto Size And Hide Column.xlsx"))) {
            Sheet sheet = reader.sheet(0);
            assert "期末成绩".equals(sheet.getName());
            List<Map<String, Object>> list = sheet.header(1, 3).rows().map(Row::toMap).collect(Collectors.toList());
            assert expectList.size() == list.size();
            for (int i = 0, len = expectList.size(); i < len; i++) {
                ListObjectSheetTest.Student expect = expectList.get(i);
                Map<String, Object> o = list.get(i);
                assert expect.getId() == Integer.parseInt(o.get("共用表头:学号").toString());
                assert expect.getName().equals(o.get("共用表头:姓名").toString());
                assert expect.getScore() == Integer.parseInt(o.get("成绩").toString());
            }
        }
    }

    @Test public void testAutoSizeAndHideColPaging() throws IOException {
        List<ListObjectSheetTest.Student> expectList = ListObjectSheetTest.Student.randomTestData();
        IWorksheetWriter worksheetWriter;
        new Workbook().setAutoSize(true)
            .addSheet(new ListSheet<>("期末成绩", expectList
                , new Column().addSubColumn(new ListSheet.EntryColumn("共用表头")).addSubColumn(new Column("学号", "id").setHeaderComment(new Comment("abc", "content")))
                , new ListSheet.EntryColumn("共用表头").addSubColumn(new Column("姓名", "name"))
                , new Column("成绩", "score").hide()
            ).setSheetWriter(worksheetWriter = new XMLWorksheetWriter() {
                @Override
                public int getRowLimit() {
                    return 10;
                }
            })).writeTo(defaultTestPath.resolve("Auto Size And Hide Column Paging.xlsx"));

        int count = expectList.size(), rowLimit = worksheetWriter.getRowLimit();
        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("Auto Size And Hide Column Paging.xlsx"))) {
            assert reader.getSize() == (count % rowLimit > 0 ? count / rowLimit + 1 : count / (rowLimit - 1)); // Include header row

            for (int i = 0, len = reader.getSize(), a = 0; i < len; i++) {
                List<Map<String, Object>> list = reader.sheet(i).header(1, 3).rows().map(Row::toMap).collect(Collectors.toList());
                if (i < len - 1) assert list.size() == rowLimit;
                else assert expectList.size() - rowLimit * (len - 1) == list.size();
                for (int j = 0; j < list.size(); j++) {
                    ListObjectSheetTest.Student expect = expectList.get(a++);
                    Map<String, Object> o = list.get(j);
                    assert expect.getId() == Integer.parseInt(o.get("共用表头:学号").toString());
                    assert expect.getName().equals(o.get("共用表头:姓名").toString());
                    assert expect.getScore() == Integer.parseInt(o.get("成绩").toString());
                }
            }
        }
    }

    @Test public void testMapRepeatHeader() throws IOException {
        List<Map<String, Object>> expectList = new ArrayList<>();
        new Workbook()
            .addSheet(new ListMapSheet("Map"
                  , new Column("aaa").addSubColumn(new Column("boolean", "bv"))
                  , new Column("aaa").addSubColumn(new Column("char", "cv"))
                  , new Column("short", "sv")
                  , new Column("int", "nv")
                  , new Column("long", "lv")
                  , new Column("LocalDateTime", "ldtv").setNumFmt(NumFmt.DATETIME_FORMAT)
                  , new Column("LocalTime", "ltv").setNumFmt(NumFmt.TIME_FORMAT)) {
                  int i = 3;

                  @Override
                  protected List<Map<String, ?>> more() {
                      List<Map<String, Object>> a = new ArrayList<>();
                      for (; i > 0; i--) {
                          Map<String, Object> data = new LinkedHashMap<>();
                          data.put("bv", random.nextInt(10) < 3);
                          data.put("cv", random.nextInt(26) + 'A');
                          data.put("sv", random.nextInt());
                          data.put("nv", random.nextInt());
                          data.put("lv", random.nextInt());
                          data.put("ldtv", LocalDateTime.now());
                          data.put("ltv", LocalTime.now());
                          a.add(data);
                      }
                      expectList.addAll(a);
                      return new ArrayList<>(a);
                  }
              }
            ).writeTo(defaultTestPath.resolve("Map Repeat Header.xlsx"));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("Map Repeat Header.xlsx"))) {
            Sheet sheet = reader.sheet(0);
            assert "Map".equals(sheet.getName());
            List<Map<String, Object>> list = sheet.header(1, 2).rows().map(Row::toMap).collect(Collectors.toList());
            assert expectList.size() == list.size();
            for (int i = 0, len = expectList.size(); i < len; i++) {
                Map<String, Object> expect = expectList.get(i), o = list.get(i);
                assert expect.get("bv").equals(o.get("aaa:boolean"));
                assert expect.get("cv").equals(o.get("aaa:char"));
                assert expect.get("sv").equals(o.get("short"));
                assert expect.get("nv").equals(o.get("int"));
                assert expect.get("lv").equals(o.get("long"));
                assert Timestamp.valueOf((LocalDateTime) expect.get("ldtv")).getTime() / 1000 == ((Timestamp) o.get("LocalDateTime")).getTime() / 1000;
                LocalTime t0 = (LocalTime) expect.get("ltv");
                Time t1 = (Time) o.get("LocalTime");
                assert t0.getHour() == t1.getHours();
                assert t0.getMinute() == t1.getMinutes();
                assert t0.getSecond() == t1.getSeconds();
            }
        }
    }

    @Test public void testRepeatColumnFromN() throws IOException {
        List<RepeatableEntry4> list = RepeatableEntry4.randomTestData();
        int startRowIndex = 7;
        new Workbook().setAutoSize(true)
            .addSheet(new ListSheet<>(list).setStartRowIndex(startRowIndex))
            .writeTo(defaultTestPath.resolve("Repeat Columns From 7.xlsx"));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("Repeat Columns From 7.xlsx"))) {
            List<RepeatableEntry4> readList = reader.sheet(0).header(startRowIndex, startRowIndex + 1).bind(RepeatableEntry4.class).rows()
                .map(row -> (RepeatableEntry4) row.get()).collect(Collectors.toList());

            assert list.size() == readList.size();
            for (int i = 0, len = list.size(); i < len; i++)
                assert list.get(i).equals(readList.get(i));

            // Row to Map
            List<Map<String, Object>> mapList = reader.sheet(0).header(startRowIndex, startRowIndex + 1).rows().map(Row::toMap).collect(Collectors.toList());
            assert list.size() == mapList.size();
            for (int i = 0, len = list.size(); i < len; i++) {
                Map<String, Object> sub = mapList.get(i);
                RepeatableEntry4 src = list.get(i);

                assert sub.get("订单号").equals(src.orderNo);
                assert sub.get("收件人").equals(src.recipient);
                assert sub.get("收件地址:省").equals(src.province);
                assert sub.get("收件地址:市").equals(src.city);
                assert sub.get("收件地址:区").equals(src.area);
                assert sub.get("收件地址:详细地址").equals(src.detail);
            }
        }
    }

    @Test public void testRepeatColumnFromStayAtA1() throws IOException {
        List<RepeatableEntry4> list = RepeatableEntry4.randomTestData();
        int startRowIndex = 7;
        new Workbook().setAutoSize(true)
            .addSheet(new ListSheet<>(list).setStartRowIndex(startRowIndex, false))
            .writeTo(defaultTestPath.resolve("Repeat Columns From 7 Stay at A1.xlsx"));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("Repeat Columns From 7 Stay at A1.xlsx"))) {
            List<RepeatableEntry4> readList = reader.sheet(0).header(startRowIndex, startRowIndex + 1).bind(RepeatableEntry4.class).rows()
                .map(row -> (RepeatableEntry4) row.get()).collect(Collectors.toList());

            assert list.size() == readList.size();
            for (int i = 0, len = list.size(); i < len; i++)
                assert list.get(i).equals(readList.get(i));

            // Row to Map
            List<Map<String, Object>> mapList = reader.sheet(0).header(startRowIndex, startRowIndex + 1).rows().map(Row::toMap).collect(Collectors.toList());
            assert list.size() == mapList.size();
            for (int i = 0, len = list.size(); i < len; i++) {
                Map<String, Object> sub = mapList.get(i);
                RepeatableEntry4 src = list.get(i);

                assert sub.get("订单号").equals(src.orderNo);
                assert sub.get("收件人").equals(src.recipient);
                assert sub.get("收件地址:省").equals(src.province);
                assert sub.get("收件地址:市").equals(src.city);
                assert sub.get("收件地址:区").equals(src.area);
                assert sub.get("收件地址:详细地址").equals(src.detail);
            }
        }
    }

    @Test public void testRepeat2AddressHeaders() throws IOException {
        List<RepeatableEntry5> list = RepeatableEntry5.randomTestData(20);
        new Workbook().setAutoSize(true)
            .addSheet(new ListSheet<>(list))
            .writeTo(defaultTestPath.resolve("Repeat 2 Address Headers.xlsx"));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("Repeat 2 Address Headers.xlsx"))) {
            List<RepeatableEntry5> readList = reader.sheet(0).header(1, 2).rows().map(row -> row.to(RepeatableEntry5.class)).collect(Collectors.toList());

            assert list.size() == readList.size();
            for (int i = 0, len = list.size(); i < len; i++)
                assert list.get(i).equals(readList.get(i));

            // Row to Map
            List<Map<String, Object>> mapList = reader.sheet(0).header(1, 2).rows().map(Row::toMap).collect(Collectors.toList());
            assert list.size() == mapList.size();
            for (int i = 0, len = list.size(); i < len; i++) {
                Map<String, Object> sub = mapList.get(i);
                RepeatableEntry5 src = list.get(i);

                assert sub.get("运单号").equals(src.orderNo);
                assert sub.get("收件地址:省").equals(src.rProvince);
                assert sub.get("收件地址:市").equals(src.rCity);
                assert sub.get("收件地址:详细地址").equals(src.rDetail);
                assert sub.get("收件人").equals(src.recipient);
                assert sub.get("寄件地址:省").equals(src.sProvince);
                assert sub.get("寄件地址:市").equals(src.sCity);
                assert sub.get("寄件地址:详细地址").equals(src.sDetail);
                assert sub.get("寄件人").equals(src.sender);
            }
        }
    }

    public static final String[] provinces = {"江苏省", "湖北省", "浙江省", "广东省"};
    public static final String[][] cities = {{"南京市", "苏州市", "无锡市", "徐州市"}
        , {"武汉市", "黄冈市", "黄石市", "孝感市", "宜昌市"}
        , {"杭州市", "温州市", "绍兴市", "嘉兴市"}
        , {"广州市", "深圳市", "佛山市"}
    };
    public static final String[][][] areas = {{
        {"玄武区", "秦淮区", "鼓楼区", "雨花台区", "栖霞区"}
        , {"虎丘区", "吴中区", "相城区", "姑苏区", "吴江区"}
        , {"锡山区", "惠山区", "滨湖区", "新吴区", "江阴市"}
        , {"鼓楼区", "云龙区", "贾汪区", "泉山区"}
    }, {
        {"江岸区", "江汉区", "硚口区", "汉阳区", "武昌区", "青山区", "洪山区", "东西湖区"}
        , {"黄州区", "团风县", "红安县"}
        , {"黄石港区", "西塞山区", "下陆区", "铁山区"}
        , {"孝南区", "孝昌县", "大悟县", "云梦县"}
        , {"西陵区", "伍家岗区", "点军区"}
    }, {
        {"上城区", "下城区", "江干区", "拱墅区", "西湖区", "滨江区", "余杭区", "萧山区"}
        , {"鹿城区", "龙湾区", "洞头区"}
        , {"越城区", "柯桥区", "上虞区", "新昌县", "诸暨市", "嵊州市"}
        , {"南湖区", "秀洲区", "嘉善县", "海盐县", "海宁市", "平湖市", "桐乡市"}
    }, {
        {"荔湾区", "白云区", "天河区", "黄埔区", "番禺区", "花都区"}
        , {"罗湖区", "福田区", "南山区", "龙岗区"}
        , {"禅城区", "南海区", "顺德区", "三水区", "高明区"}
    }};

    public static class RepeatableEntry {
        @ExcelColumn("TOP")
        @ExcelColumn("K")
        @ExcelColumn
        @ExcelColumn("订单号")
        private String orderNo;
        @ExcelColumn("TOP")
        @ExcelColumn("K")
        @ExcelColumn("A")
        @ExcelColumn("收件人")
        private String recipient;
        @ExcelColumn("TOP")
        @ExcelColumn("收件地址")
        @ExcelColumn("A")
        @ExcelColumn("省")
        private String province;
        @ExcelColumn("TOP")
        @ExcelColumn("收件地址")
        @ExcelColumn("A")
        @ExcelColumn("市")
        private String city;
        @ExcelColumn("TOP")
        @ExcelColumn("收件地址")
        @ExcelColumn("B")
        @ExcelColumn("区")
        private String area;
        @ExcelColumn("TOP")
        @ExcelColumn(value = "收件地址", comment = @HeaderComment("精确到门牌号"))
        @ExcelColumn("B")
        @ExcelColumn("详细地址")
        private String detail;

        public RepeatableEntry() {}

        public RepeatableEntry(String orderNo, String recipient, String province, String city, String area, String detail) {
            this.orderNo = orderNo;
            this.recipient = recipient;
            this.province = province;
            this.city = city;
            this.area = area;
            this.detail = detail;
        }

        public static List<RepeatableEntry> randomTestData(int n) {
            List<RepeatableEntry> list = new ArrayList<>(n);
            for (int i = 0, p, c; i < n; i++) {
                list.add(new RepeatableEntry(Integer.toString(Math.abs(random.nextInt())), getRandomString(8) + 2, provinces[p = random.nextInt(provinces.length)], cities[p][c = random.nextInt(cities[p].length)], areas[p][c][random.nextInt(areas[p][c].length)], "xx街" + (random.nextInt(10) + 1) + "号"));
            }
            return list;
        }

        public static List<RepeatableEntry> randomTestData() {
            int n = random.nextInt(100) + 1;
            return randomTestData(n);
        }

        public String getOrderNo() {
            return orderNo;
        }

        public String getRecipient() {
            return recipient;
        }

        public String getProvince() {
            return province;
        }

        public String getCity() {
            return city;
        }

        public String getArea() {
            return area;
        }

        public String getDetail() {
            return detail;
        }

        @Override
        public int hashCode() {
            return orderNo.hashCode();
        }

        @Override
        public boolean equals(Object o) {
            if (o instanceof RepeatableEntry) {
                RepeatableEntry other = (RepeatableEntry) o;
                return Objects.equals(orderNo, other.orderNo)
                    && Objects.equals(recipient, other.recipient)
                    && Objects.equals(province, other.province)
                    && Objects.equals(city, other.city)
                    && Objects.equals(detail, other.detail);
            }
            return false;
        }

        @Override
        public String toString() {
            return orderNo + " | " + recipient + " | " + province + " | " + city + " | " + area + " | " + detail;
        }
    }

    public static class RepeatableEntry3 {
        @ExcelColumn("TOP")
        @ExcelColumn("K")
        @ExcelColumn
        @ExcelColumn("订单号")
        private String orderNo;
        @ExcelColumn("TOP")
        @ExcelColumn("K")
        @ExcelColumn("A")
        @ExcelColumn("收件人")
        private String recipient;
        @ExcelColumn("TOP")
        @ExcelColumn("收件地址")
        @ExcelColumn("A")
        @ExcelColumn("省")
        private String province;
        @ExcelColumn("TOP")
        @ExcelColumn("市")
        @ExcelColumn("市")
        @ExcelColumn("市")
        private String city;
        @ExcelColumn("TOP")
        @ExcelColumn("收件地址")
        @ExcelColumn("B")
        @ExcelColumn("区")
        private String area;
        @ExcelColumn("详细地址")
        @ExcelColumn(value = "详细地址", comment = @HeaderComment("精确到门牌号"))
        @ExcelColumn("详细地址")
        @ExcelColumn("详细地址")
        private String detail;

        public RepeatableEntry3() {}

        public RepeatableEntry3(String orderNo, String recipient, String province, String city, String area, String detail) {
            this.orderNo = orderNo;
            this.recipient = recipient;
            this.province = province;
            this.city = city;
            this.area = area;
            this.detail = detail;
        }

        public static List<RepeatableEntry3> randomTestData(int n) {
            List<RepeatableEntry3> list = new ArrayList<>(n);
            for (int i = 0, p, c; i < n; i++) {
                list.add(new RepeatableEntry3(Integer.toString(Math.abs(random.nextInt())), getRandomString(8) + 2, provinces[p = random.nextInt(provinces.length)], cities[p][c = random.nextInt(cities[p].length)], areas[p][c][random.nextInt(areas[p][c].length)], "xx街" + (random.nextInt(10) + 1) + "号"));
            }
            return list;
        }

        public static List<RepeatableEntry3> randomTestData() {
            int n = random.nextInt(100) + 1;
            return randomTestData(n);
        }

        @Override
        public int hashCode() {
            return orderNo.hashCode();
        }

        @Override
        public boolean equals(Object o) {
            if (o instanceof RepeatableEntry3) {
                RepeatableEntry3 other = (RepeatableEntry3) o;
                return Objects.equals(orderNo, other.orderNo)
                    && Objects.equals(recipient, other.recipient)
                    && Objects.equals(province, other.province)
                    && Objects.equals(city, other.city)
                    && Objects.equals(detail, other.detail);
            }
            return false;
        }

        @Override
        public String toString() {
            return orderNo + " | " + recipient + " | " + province + " | " + city + " | " + area + " | " + detail;
        }
    }

    public static class RepeatableEntry4 {
        @ExcelColumn(value = "订单号", colIndex = 3)
        private String orderNo;
        @ExcelColumn(value = "收件人", colIndex = 4)
        private String recipient;
        @ExcelColumn("收件地址")
        @ExcelColumn(value = "省", colIndex = 5)
        private String province;
        @ExcelColumn("收件地址")
        @ExcelColumn(value = "市", colIndex = 6)
        private String city;
        @ExcelColumn("收件地址")
        @ExcelColumn(value = "区", colIndex = 7)
        private String area;
        @ExcelColumn("收件地址")
        @ExcelColumn(value = "详细地址", colIndex = 8)
        private String detail;

        public RepeatableEntry4() {}

        public RepeatableEntry4(String orderNo, String recipient, String province, String city, String area, String detail) {
            this.orderNo = orderNo;
            this.recipient = recipient;
            this.province = province;
            this.city = city;
            this.area = area;
            this.detail = detail;
        }

        public static List<RepeatableEntry4> randomTestData(int n) {
            List<RepeatableEntry4> list = new ArrayList<>(n);
            for (int i = 0, p, c; i < n; i++) {
                list.add(new RepeatableEntry4(Integer.toString(Math.abs(random.nextInt())), getRandomString(8) + 2, provinces[p = random.nextInt(provinces.length)], cities[p][c = random.nextInt(cities[p].length)], areas[p][c][random.nextInt(areas[p][c].length)], "xx街" + (random.nextInt(10) + 1) + "号"));
            }
            return list;
        }

        public static List<RepeatableEntry4> randomTestData() {
            int n = random.nextInt(100) + 1;
            return randomTestData(n);
        }

        public String getOrderNo() {
            return orderNo;
        }

        public String getRecipient() {
            return recipient;
        }

        public String getProvince() {
            return province;
        }

        public String getCity() {
            return city;
        }

        public String getArea() {
            return area;
        }

        public String getDetail() {
            return detail;
        }

        @Override
        public int hashCode() {
            return orderNo.hashCode();
        }

        @Override
        public boolean equals(Object o) {
            if (o instanceof RepeatableEntry4) {
                RepeatableEntry4 other = (RepeatableEntry4) o;
                return Objects.equals(orderNo, other.orderNo)
                    && Objects.equals(recipient, other.recipient)
                    && Objects.equals(province, other.province)
                    && Objects.equals(city, other.city)
                    && Objects.equals(detail, other.detail);
            }
            return false;
        }

        @Override
        public String toString() {
            return orderNo + " | " + recipient + " | " + province + " | " + city + " | " + area + " | " + detail;
        }
    }

    public static class RepeatableEntry5 {
    @ExcelColumn("运单号")
    private String orderNo;
    @ExcelColumn("收件地址")
    @ExcelColumn("省")
    private String rProvince;
    @ExcelColumn("收件地址")
    @ExcelColumn("市")
    private String rCity;
    @ExcelColumn("收件地址")
    @ExcelColumn("详细地址")
    private String rDetail;
    @ExcelColumn("收件人")
    private String recipient;
    @ExcelColumn("寄件地址")
    @ExcelColumn("省")
    private String sProvince;
    @ExcelColumn("寄件地址")
    @ExcelColumn("市")
    private String sCity;
    @ExcelColumn("寄件地址")
    @ExcelColumn("详细地址")
    private String sDetail;
    @ExcelColumn("寄件人")
    private String sender;

        public RepeatableEntry5() { }

        public RepeatableEntry5(String orderNo, String rProvince, String rCity, String rDetail, String recipient, String sProvince, String sCity, String sDetail, String sender) {
            this.orderNo = orderNo;
            this.rProvince = rProvince;
            this.rCity = rCity;
            this.rDetail = rDetail;
            this.recipient = recipient;
            this.sProvince = sProvince;
            this.sCity = sCity;
            this.sDetail = sDetail;
            this.sender = sender;
        }

        public static List<RepeatableEntry5> randomTestData(int n) {
            List<RepeatableEntry5> list = new ArrayList<>(n);
            for (int i = 0, p; i < n; i++) {
                list.add(new RepeatableEntry5(Integer.toString(Math.abs(random.nextInt())), provinces[p = random.nextInt(provinces.length)], cities[p][random.nextInt(cities[p].length)], "xx街" + (random.nextInt(10) + 1) + "号", "王**", provinces[p = random.nextInt(provinces.length)], cities[p][random.nextInt(cities[p].length)], "xx街" + (random.nextInt(10) + 1) + "号", "周**"));
            }
            return list;
        }

        @Override
        public int hashCode() {
            return orderNo.hashCode();
        }

        @Override
        public boolean equals(Object o) {
            if (o instanceof RepeatableEntry5) {
                RepeatableEntry5 other = (RepeatableEntry5) o;
                return Objects.equals(orderNo, other.orderNo)
                    && Objects.equals(rProvince, other.rProvince)
                    && Objects.equals(rCity, other.rCity)
                    && Objects.equals(rDetail, other.rDetail)
                    && Objects.equals(recipient, other.recipient)
                    && Objects.equals(sProvince, other.sProvince)
                    && Objects.equals(sCity, other.sCity)
                    && Objects.equals(sDetail, other.sDetail)
                    && Objects.equals(sender, other.sender);
            }
            return false;
        }
    }

}

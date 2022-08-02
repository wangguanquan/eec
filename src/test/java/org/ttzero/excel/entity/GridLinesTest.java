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
import org.ttzero.excel.annotation.HeaderStyle;
import org.ttzero.excel.entity.e7.XMLWorksheetWriter;

import java.io.IOException;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Timestamp;
import java.util.List;
import java.util.function.Supplier;

import static org.ttzero.excel.entity.ListMapSheetTest.createTestData;

/**
 * @author guanquan.wang at 2022-08-02 19:17
 */
public class GridLinesTest extends SQLWorkbookTest {
    @Test public void testListSheet() throws IOException {
        new Workbook("ListSheet ignore grid lines")
            .addSheet(new ListSheet<>(ListObjectSheetTest.Item.randomTestData()).hideGridLines())
            .writeTo(defaultTestPath);
    }

    @Test public void testListSheetAnnotation() throws IOException {
        new Workbook("ListSheet annotation ignore grid lines")
            .setAutoSize(true)
            .addSheet(new ListSheet<>(HideGridLineAllType.randomTestData()).hideGridLines())
            .writeTo(defaultTestPath);
    }

    @Test public void testListSheetPaging() throws IOException {
        new Workbook("ListSheet Paging ignore grid lines")
            .setAutoSize(true)
            .addSheet(new ListSheet<>(ListObjectSheetTest.Item.randomTestData()).hideGridLines()
                .setSheetWriter(new XMLWorksheetWriter() {
                @Override
                public int getRowLimit() {
                    return 10;
                }
            }))
            .writeTo(defaultTestPath);
    }

    @Test public void testListMapSheet() throws IOException {
        new Workbook("ListMapSheet ignore grid lines")
            .addSheet(new ListMapSheet(createTestData()).hideGridLines())
            .writeTo(defaultTestPath);
    }

    @Test public void testStatementSheet() throws SQLException, IOException {
        try (Connection con = getConnection()) {
            new Workbook("StatementSheet ignore grid lines")
                .addSheet(new StatementSheet(con, "select id, name, age, create_date, update_date from student order by age"
                    , new Column("学号", int.class)
                    , new Column("性名", String.class)
                    , new Column("年龄", int.class)
                    , new Column("创建时间", Timestamp.class).setColIndex(0)
                    , new Column("更新", Timestamp.class)
                ).hideGridLines()).writeTo(defaultTestPath);
        }
    }

    @Test public void testResultSetSheet() throws SQLException, IOException {
        try (
            Connection con = getConnection();
            PreparedStatement ps = con.prepareStatement("select id, name, age, create_date, update_date from student limit 10");
            ResultSet rs = ps.executeQuery()
        ) {
            new Workbook("ResultSetSheet ignore grid lines")
                .setConnection(con)
                .addSheet(new ResultSetSheet(rs
                    , new Column("学号", int.class)
                    , new Column("性名", String.class)
                    , new Column("年龄", Integer.class)
                    , new Column("创建时间", Timestamp.class)
                    , new Column("更新", Timestamp.class)
                ).hideGridLines()).writeTo(defaultTestPath);
        }
    }

    @HeaderStyle(showGridLines = false)
    public static class HideGridLineAllType extends ListObjectSheetTest.AllType {
        public static List<ListObjectSheetTest.AllType> randomTestData() {
            return randomTestData(HideGridLineAllType::new);
        }

        public static List<ListObjectSheetTest.AllType> randomTestData(Supplier<ListObjectSheetTest.AllType> sup) {
            int size = random.nextInt(100) + 1;
            return randomTestData(size, sup);
        }
    }
}

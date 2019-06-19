/*
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
import org.ttzero.excel.Print;

import java.io.IOException;
import java.sql.Connection;
import java.sql.SQLException;

/**
 * Create by guanquan.wang at 2019-04-28 22:47
 */
public class StatementPagingTest extends SQLWorkbookTest {
    @Test public void testPaging() {
        try (Connection con = getConnection()) {
            new Workbook("statement paging", author)
                .watch(Print::println)
                .setConnection(con)
                .addSheet("select id, name, age from student"
                    , new Sheet.Column("学号", int.class)
                    , new Sheet.Column("性名", String.class)
                    , new Sheet.Column("年龄", int.class)
                )
                .setWorkbookWriter(new ReLimitXMLWorkbookWriter())
                .writeTo(defaultTestPath);
        } catch (SQLException |IOException e) {
            e.printStackTrace();
        }
    }
}

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

package cn.ttzero.excel.entity;

import cn.ttzero.excel.Print;
import cn.ttzero.excel.entity.style.Fill;
import cn.ttzero.excel.entity.style.PatternType;
import cn.ttzero.excel.entity.style.Styles;
import org.junit.Test;

import java.awt.*;
import java.io.IOException;
import java.sql.Connection;
import java.sql.SQLException;

/**
 * Create by guanquan.wang at 2019-04-28 22:47
 */
public class StatementSheetTest extends SQLWorkbookTest {
    @Test public void testWrite() {
        try (Connection con = getConnection()) {
            new Workbook("statement", author)
                .watch(Print::println)
                .setConnection(con)
                .addSheet("select id, name, age from student order by age"
                    , new Sheet.Column("学号", int.class)
                    , new Sheet.Column("性名", String.class)
                    , new Sheet.Column("年龄", int.class)
                )
                .writeTo(defaultTestPath);
        } catch (SQLException |IOException e) {
            e.printStackTrace();
        }
    }


    @Test public void testStyleProcessor() {
        try (Connection con = getConnection()) {
            new Workbook("statement style processor", author)
                    .watch(Print::println)
                    .setConnection(con)
                    .addSheet("select id, name, age from student"
                            , new Sheet.Column("学号", int.class)
                            , new Sheet.Column("性名", String.class)
                            , new Sheet.Column("年龄", int.class)
                                .setStyleProcessor((o, style, sst) -> {
                                    int n = (int) o;
                                    if (n < 10) {
                                        style = Styles.clearFill(style)
                                                | sst.addFill(new Fill(PatternType.solid, Color.orange));
                                    }
                                    return style;
                                })
                    )
                    .writeTo(defaultTestPath);
        } catch (SQLException |IOException e) {
            e.printStackTrace();
        }
    }
}

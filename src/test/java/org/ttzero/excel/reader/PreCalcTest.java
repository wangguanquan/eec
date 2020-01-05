/*
 * Copyright (c) 2019-2021, guanquan.wang@yandex.com All Rights Reserved.
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

package org.ttzero.excel.reader;

import org.junit.Test;

import static org.ttzero.excel.Print.println;

/**
 * Create by guanquan.wang at 2020-01-05 20:35
 */
public class PreCalcTest {
    @Test public void testSimple() {
        Row row = new MyRow();
        row.addRef(0, "B2:B8");
        row.setCalc(0, "A89");
        String calc = row.getCalc(0, 5 << 14 | 2);
        assert "A92".equals(calc);
    }

    @Test public void testPlus() {
        Row row = new MyRow();
        row.addRef(0, "B2:B8");
        row.setCalc(0, "(A2+A3)+1");
        String calc = row.getCalc(0, 4 << 14 | 2);
        println(calc);
        assert "(A4+A5)+1".equals(calc);
    }

    @Test public void testRange() {
        Row row = new MyRow();
        row.addRef(0, "B2:D8");
        row.setCalc(0, "(A2+B3)+1");
        String calc = row.getCalc(0, 4 << 14 | 3);
        println(calc);
        assert "(B4+C5)+1".equals(calc);
    }

    @Test public void testDoubleQuotes() {
        Row row = new MyRow();
        row.addRef(0, "B2:D8");
        row.setCalc(0, "\"AB12\".substring(3)");
        String calc = row.getCalc(0, 4 << 14 | 3);
        println(calc);
        assert "\"AB12\".substring(3)".equals(calc);
    }

    @Test public void testDoubleQuotes2() {
        Row row = new MyRow();
        row.addRef(0, "B2:D8");
        row.setCalc(0, "\"AB12.substring(3)");
        String calc = row.getCalc(0, 4 << 14 | 3);
        println(calc);
        assert "\"AB12.substring(3)".equals(calc);
    }

    @Test public void testRows() {
        Row row = new MyRow();
        row.addRef(1, "B9:B65");
        row.setCalc(1, "A9+1");
        String calc = row.getCalc(1, 10 << 14 | 2);
        println(calc);
        assert "A10+1".equals(calc);
    }

    private static class MyRow extends Row { }
}

/*
 * Copyright (c) 2017-2021, guanquan.wang@yandex.com All Rights Reserved.
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

import static org.junit.Assert.assertEquals;

/**
 * @author guanquan.wang at 2021-03-25 00:54
 */
public class DimensionTest {

    @Test public void testFirstDim() {
        Dimension d = Dimension.of("A1");
        assertEquals(d.firstColumn, 1);
        assertEquals(d.firstRow, 1);
        assertEquals(d.lastColumn, 1);
        assertEquals(d.lastRow, 1);
        assertEquals(d.width, 1);
        assertEquals(d.height, 1);

        assertEquals(d.toReferer(), "$A$1");
    }

    @Test public void testFirstDim2() {
        Dimension d = Dimension.of("B3:");
        assertEquals(d.firstColumn, 2);
        assertEquals(d.firstRow, 3);
        assertEquals(d.lastColumn, 2);
        assertEquals(d.lastRow, 3);
        assertEquals(d.width, 1);
        assertEquals(d.height, 1);

        assertEquals(d.toReferer(), "$B$3");
    }

    @Test public void testLastDim() {
        Dimension d = Dimension.of(":C2");
        assertEquals(d.firstColumn, 1);
        assertEquals(d.firstRow, 1);
        assertEquals(d.lastColumn, 3);
        assertEquals(d.lastRow, 2);
        assertEquals(d.width, 3);
        assertEquals(d.height, 2);

        assertEquals(d.toReferer(), "$A$1:$C$2");
    }

    @Test public void testFullDim() {
        Dimension d = Dimension.of("AZ103:CCA63335");
        assertEquals(d.firstColumn, 52);
        assertEquals(d.firstRow, 103);
        assertEquals(d.lastColumn, 2107);
        assertEquals(d.lastRow, 63335);
        assertEquals(d.width, 2056);
        assertEquals(d.height, 63233);

        assertEquals(d.toReferer(), "$AZ$103:$CCA$63335");
    }
}

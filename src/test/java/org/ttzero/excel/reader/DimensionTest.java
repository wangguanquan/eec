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

/**
 * @author guanquan.wang at 2021-03-25 00:54
 */
public class DimensionTest {

    @Test public void testFirstDim() {
        Dimension d = Dimension.of("A1");
        assert d.firstColumn == 1;
        assert d.firstRow == 1;
        assert d.lastColumn == 1;
        assert d.lastRow == 1;
        assert d.width == 1;
        assert d.height == 1;
    }

    @Test public void testFirstDim2() {
        Dimension d = Dimension.of("B3:");
        assert d.firstColumn == 2;
        assert d.firstRow == 3;
        assert d.lastColumn == 2;
        assert d.lastRow == 3;
        assert d.width == 1;
        assert d.height == 1;
    }

    @Test public void testLastDim() {
        Dimension d = Dimension.of(":C2");
        assert d.firstColumn == 1;
        assert d.firstRow == 1;
        assert d.lastColumn == 3;
        assert d.lastRow == 2;
        assert d.width == 3;
        assert d.height == 2;
    }

    @Test public void testFullDim() {
        Dimension d = Dimension.of("A1:C2");
        assert d.firstColumn == 1;
        assert d.firstRow == 1;
        assert d.lastColumn == 3;
        assert d.lastRow == 2;
        assert d.width == 3;
        assert d.height == 2;
    }
}

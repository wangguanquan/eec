/*
 * Copyright (c) 2017-2020, guanquan.wang@yandex.com All Rights Reserved.
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
import static org.ttzero.excel.reader.Grid.FastGrid.isPowerOfTwo;

/**
 * @author guanquan.wang at 2020-01-09 17:19
 */
public class GridTest {
    @Test public void testGrid1() {
        Grid grid = GridFactory.create(new Dimension(1, (short) 1, 10, (short) 1));

        grid.mark(new Dimension(3, (short) 1, 7, (short) 1));

        println(grid);

        assert !grid.test(2, 1);
        assert grid.test(3, 1);
        assert grid.test(7, 1);
        assert !grid.test(8, 1);
    }

    @Test public void testGrid2() {
        Grid grid = GridFactory.create(new Dimension(1, (short) 1, 10, (short) 2));

        grid.mark(new Dimension(3, (short) 1, 7, (short) 1));

        println(grid);

        assert !grid.test(2, 1);
        assert grid.test(3, 1);
        assert grid.test(7, 1);
        assert !grid.test(8, 1);
    }

    @Test public void testGrid4() {
        Grid grid = GridFactory.create(new Dimension(1, (short) 1, 10, (short) 3));

        grid.mark(new Dimension(3, (short) 1, 7, (short) 3));

        println(grid);

        assert !grid.test(2, 2);
        assert grid.test(3, 2);
        assert grid.test(3, 3);
        assert grid.test(7, 1);
        assert grid.test(7, 2);
        assert !grid.test(8, 2);
    }

    @Test public void testGrid8() {
        Grid grid = GridFactory.create(new Dimension(1, (short) 1, 10, (short) 8));

        grid.mark(new Dimension(3, (short) 4, 7, (short) 4));

        println(grid);
    }

    @Test public void testGrid16() {
        Grid grid = GridFactory.create(new Dimension(1, (short) 1, 10, (short) 10));

        grid.mark(new Dimension(4, (short) 5, 9, (short) 7));

        println(grid);
    }

    @Test public void testGrid162() {
        Grid grid = GridFactory.create(new Dimension(1, (short) 1, 10, (short) 10));

        grid.mark(new Dimension(2, (short) 2, 4, (short) 6));
        grid.mark(new Dimension(3, (short) 7, 5, (short) 9));
        grid.mark(new Dimension(7, (short) 10, 10, (short) 10));

        println(grid);

        assert grid.test(7,10);
        assert !grid.test(6, 10);
        assert !grid.test(7, 9);
        assert grid.test(5, 8);
    }

    @Test public void testGrid32() {
        Grid grid = GridFactory.create(Dimension.of("A1:AF10"));

        grid.mark(Dimension.of("G3:AA9"));

        println(grid);
    }

    @Test public void testGrid64() {
        Grid grid = GridFactory.create(new Dimension(1, (short) 1, 10, (short) 54));

        grid.mark(new Dimension(2, (short) 2, 4, (short) 6));

        println(grid);
    }

    @Test public void testPowerOfTwo() {
        assert isPowerOfTwo(1);
        assert isPowerOfTwo(2);
        assert isPowerOfTwo(1024);
        assert !isPowerOfTwo(3);
        assert !isPowerOfTwo(6);
    }
}

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

import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.List;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertTrue;
import static org.ttzero.excel.reader.Grid.FastGrid.isPowerOfTwo;

/**
 * @author guanquan.wang at 2020-01-09 17:19
 */
public class GridTest {
    @Test public void testGridType() {
        Grid grid = GridFactory.create(Collections.singletonList(Dimension.of("A1:BM10")));
        assertTrue(grid instanceof Grid.IndexGrid);

        grid = GridFactory.create(Collections.singletonList(Dimension.of("A1:B16383")));
        assertTrue(grid instanceof Grid.FastGrid);

        grid = GridFactory.create(Collections.singletonList(Dimension.of("A1:B16384")));
        assertTrue(grid instanceof Grid.FastGrid);

    }

    @Test public void testGrid1() {
        Grid grid = GridFactory.create(Collections.singletonList(new Dimension(3, (short) 1, 7, (short) 1)));

        assertFalse(grid.test(2, 1));
        assertTrue(grid.test(3, 1));
        assertTrue(grid.test(7, 1));
        assertFalse(grid.test(8, 1));
    }

    @Test public void testGrid4() {
        Grid grid = GridFactory.create(Collections.singletonList(new Dimension(3, (short) 1, 7, (short) 3)));

        assertFalse(grid.test(2, 2));
        assertTrue(grid.test(3, 2));
        assertTrue(grid.test(3, 3));
        assertTrue(grid.test(7, 1));
        assertTrue(grid.test(7, 2));
        assertFalse(grid.test(8, 2));
    }

    @Test public void testGrid8() {
        Grid grid = GridFactory.create(Collections.singletonList(new Dimension(3, (short) 4, 7, (short) 4)));

        assertTrue(grid.test(3, 4));
        assertTrue(grid.test(4, 4));
        assertTrue(grid.test(5, 4));
        assertFalse(grid.test(5, 3));
        assertTrue(grid.test(6, 4));
        assertTrue(grid.test(7, 4));
        assertFalse(grid.test(7, 5));
        assertFalse(grid.test(8, 4));
    }

    @Test public void testGrid8_2() {
        List<Dimension> list = new ArrayList<>();
        list.add(Dimension.of("C10:D10"));
        list.add(Dimension.of("C5:D5"));
        list.add(Dimension.of("C6:D6"));
        list.add(Dimension.of("C7:D7"));
        list.add(Dimension.of("C8:D8"));
        list.add(Dimension.of("C9:D9"));
        list.add(Dimension.of("A39:A71"));
        list.add(Dimension.of("D1:E1"));
        list.add(Dimension.of("A1:A26"));
        list.add(Dimension.of("A27:A38"));
        list.add(Dimension.of("E20:H20"));
        list.add(Dimension.of("E21:H21"));
        list.add(Dimension.of("E22:H22"));
        list.add(Dimension.of("E23:H23"));
        list.add(Dimension.of("E24:H24"));
        list.add(Dimension.of("E25:H25"));
        list.add(Dimension.of("C11:D11"));
        list.add(Dimension.of("C12:D12"));
        list.add(Dimension.of("C13:D13"));
        list.add(Dimension.of("C14:D14"));
        list.add(Dimension.of("C15:D15"));
        list.add(Dimension.of("C16:D16"));

        Grid grid = GridFactory.create(list);

        assertEquals(grid.toString(), "FastGrid Size: 72B\n" +
            "00011001\n00000001\n00000001\n00000001\n" +
            "00001101\n00001101\n00001101\n00001101\n" +
            "00001101\n00001101\n00001101\n00001101\n" +
            "00001101\n00001101\n00001101\n00001101\n" +
            "00000001\n00000001\n00000001\n11110001\n" +
            "11110001\n11110001\n11110001\n11110001\n" +
            "11110001\n00000001\n00000001\n00000001\n" +
            "00000001\n00000001\n00000001\n00000001\n" +
            "00000001\n00000001\n00000001\n00000001\n" +
            "00000001\n00000001\n00000001\n00000001\n" +
            "00000001\n00000001\n00000001\n00000001\n" +
            "00000001\n00000001\n00000001\n00000001\n" +
            "00000001\n00000001\n00000001\n00000001\n" +
            "00000001\n00000001\n00000001\n00000001\n" +
            "00000001\n00000001\n00000001\n00000001\n" +
            "00000001\n00000001\n00000001\n00000001\n" +
            "00000001\n00000001\n00000001\n00000001\n" +
            "00000001\n00000001\n00000001");
    }

    @Test public void testGrid16() {
        Grid grid = GridFactory.create(Collections.singletonList(new Dimension(4, (short) 5, 9, (short) 7)));

        assertTrue(grid.test(4, 5));
        assertTrue(grid.test(4, 7));
        assertFalse(grid.test(7, 8));
        assertTrue(grid.test(8, 7));
        assertTrue(grid.test(9, 5));
        assertFalse(grid.test(9, 10));
    }

    @Test public void testGrid162() {
        List<Dimension> list = Arrays.asList(new Dimension(2, (short) 2, 4, (short) 6)
            , new Dimension(3, (short) 7, 5, (short) 9)
            , new Dimension(7, (short) 10, 10, (short) 10));
        Grid grid = GridFactory.create(list);

        assertTrue(grid.test(7, 10));
        assertFalse(grid.test(6, 10));
        assertFalse(grid.test(7, 9));
        assertTrue(grid.test(5, 8));
    }

    @Test public void testGrid32() {
        Grid grid = GridFactory.create(Collections.singletonList(Dimension.of("G3:AA9")));

        assertEquals("FastGrid Size: 32B\n" +
            "00000000000111111111111111111111\n" +
            "00000000000111111111111111111111\n" +
            "00000000000111111111111111111111\n" +
            "00000000000111111111111111111111\n" +
            "00000000000111111111111111111111\n" +
            "00000000000111111111111111111111\n" +
            "00000000000111111111111111111111", grid.toString());
    }

    @Test public void testPowerOfTwo() {
        assertTrue(isPowerOfTwo(1));
        assertTrue(isPowerOfTwo(2));
        assertTrue(isPowerOfTwo(1024));
        assertFalse(isPowerOfTwo(3));
        assertFalse(isPowerOfTwo(6));
    }

    @Test public void testLinkedScanner() {
        Grid.Scanner scanner = new Grid.LinkedScanner();
        scanner.put(new Grid.LinkedScanner.E(Dimension.of("E5:F8"), null));
        scanner.put(new Grid.LinkedScanner.E(Dimension.of("D2:F2"), null));
        scanner.put(new Grid.LinkedScanner.E(Dimension.of("B16:E17"), null));
        scanner.put(new Grid.LinkedScanner.E(Dimension.of("B2:C2"), null));
        scanner.put(new Grid.LinkedScanner.E(Dimension.of("A13:A20"), null));

        assertEquals("B2:C2->D2:F2->E5:F8->A13:A20->B16:E17", scanner.toString());

        scanner.get(5, 5);
        assertEquals("E5:F8->B2:C2->D2:F2->A13:A20->B16:E17", scanner.toString());

        scanner.get(5, 6);
        scanner.get(6, 5);
        scanner.get(6, 6);
        scanner.get(7, 5);
        scanner.get(7, 6);
        scanner.get(8, 5);
        scanner.get(8, 6);

        assertEquals("B2:C2->D2:F2->A13:A20->B16:E17->E5:F8", scanner.toString());
    }

    @Test public void testIndexGrid() {
        Dimension range = new Dimension(1, (short)1, 2, (short)17);
        List<Dimension> list = Arrays.asList(Dimension.of("H1:I1"), Dimension.of("J1:K1")
            , Dimension.of("L1:M1"), Dimension.of("N1:O1"), Dimension.of("P1:Q1"), Dimension.of("R1:S1")
            , Dimension.of("T1:U1"), Dimension.of("V1:W1"), Dimension.of("X1:Y1"), Dimension.of("Z1:AA1")
            , Dimension.of("A1:A2"), Dimension.of("B1:B2"), Dimension.of("C1:C2"), Dimension.of("D1:D2")
            , Dimension.of("E1:E2"), Dimension.of("F1:F2"), Dimension.of("G1:G2")
        );

        Grid grid = new Grid.IndexGrid(range, 2 * 17);
        for (Dimension dim : list) grid.mark(dim);

        Cell c = new Cell((short) 1);
        assertEquals(grid.merge(1, c), 1);
        assertEquals(grid.merge(2, c), 2);
        c.i = 7;
        assertEquals(grid.merge(1, c), 1);
        assertEquals(grid.merge(2, c), 2);

        c.i = 8;
        assertEquals(grid.merge(1, c), 1);
        c.i = 9;
        assertEquals(grid.merge(1, c), 2);

        c.i = 26;
        assertEquals(grid.merge(1, c), 1);
        c.i = 27;
        assertEquals(grid.merge(1, c), 2);
    }

    @Test public void testFractureGrid() {
        Grid grid = GridFactory.create(Collections.singletonList(Dimension.of("B1:C3")));

        assertFalse(grid.test(1, 1));
        assertTrue(grid.test(1, 2));
        assertTrue(grid.test(2, 2));
        assertTrue(grid.test(3, 3));
        assertFalse(grid.test(4, 2));
        assertFalse(grid.test(3, 4));
    }
}

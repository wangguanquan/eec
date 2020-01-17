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

import java.util.Arrays;
import java.util.StringJoiner;

import static java.lang.Integer.numberOfTrailingZeros;

/**
 * @author guanquan.wang at 2020-01-09 16:54
 */
interface Grid {

    /**
     * Mark `1` at the specified coordinates
     *
     * @param coordinate the excel coordinate string,
     *                   it's a coordinate or range coordinates like `A1` or `A1:C4`
     */
    default void mark(String coordinate) {
        mark(coordinate.toCharArray(), 0, coordinate.length());
    }

    /**
     * Mark `1` at the specified coordinates
     *
     * @param chars the excel coordinate buffer,
     *              it's a coordinate or range coordinates like `A1` or `A1:C4`
     * @param from  the begin index
     * @param to    the end index
     */
    void mark(char[] chars, int from, int to);

    /**
     * Mark `1` at the specified {@link Dimension}
     *
     * @param dimension range {@link Dimension}
     */
    void mark(Dimension dimension);

    boolean test(int r, int c);
}

class GridFactory {
    private GridFactory() { }

    static Grid create(Dimension dim) {
        int r = dim.lastRow - dim.firstRow + 1
            , c = dim.lastColumn - dim.firstColumn + 1;

        return c <= 128 && r < 1 << 14 ? new FastGrid(dim) : new FractureGrid();
    }
}

class FastGrid implements Grid {
    private int fr, fc, lr, lc; // Start index of Row and Column(One base)
    private long[] g;

    private int b // Power of two minus 1
        , c;
    FastGrid(Dimension dim) {
        fr = dim.firstRow;
        lr = dim.lastRow;
        fc = dim.firstColumn;
        lc = dim.lastColumn;

        b = powerOneBit(lc - fc + 1);
        c = numberOfTrailingZeros(b + 1) + (isPowerOfTwo(lc - fc + 1) ? -1 : 0);
        int n = 6 - c, len = (lr - fr + 1) >> n;
        g = new long[len != (len >> n << n) ? (lr - fr + 1 >> n) + 1 : (lr - fr + 1) >> n];
    }

    static int powerOneBit(int i) {
        i |= (i >>  1);
        i |= (i >>  2);
        i |= (i >>  4);
        i |= (i >>  8);
        i |= (i >> 16);
        return i;
    }

    static boolean isPowerOfTwo(int n) {
        return (n & -n) == n;
    }

    @Override
    public void mark(char[] chars, int from, int to) {

    }

    @Override
    public void mark(Dimension dimension) {
        int n = dimension.lastColumn - dimension.firstColumn + 1;
        for (int i = dimension.firstRow; i <= dimension.lastRow; i++) {
            long l = ~(~0 >> n << n);
            g[getRow(i - fr)] |= l;
        }
    }

    @Override
    public boolean test(int r, int c) {
//        return range(r, c) && g[getRow(r - fr)];
        return false;
    }

    boolean range(int r, int c) {
        return r >= fr && r <= lr && c >= fc && c <= lc;
    }

    int getRow(int i) {
        return i >> (6 - c);
    }

    @Override
    public String toString() {
        StringJoiner joiner = new StringJoiner("\n");
        for (long l : g) {
            joiner.add(append(Long.toBinaryString(l)));
        }
        return joiner.toString();
    }

    private char[] chars = new char[64];
    private String append(String s) {
        int n = s.length();
        s.getChars(0, n, chars, chars.length - n);
        Arrays.fill(chars, 0, chars.length - n, '0');
        return new String(chars);
    }
}

class FractureGrid implements Grid {

    @Override
    public void mark(char[] chars, int from, int to) {

    }

    @Override
    public void mark(Dimension dimension) {

    }

    @Override
    public boolean test(int r, int c) {
        return false;
    }
}
//
//class LongLongGrid extends FastGrid {
//
//    LongLongGrid(Dimension dim) {
//        super(dim);
//    }
//
//    @Override
//    int lg() {
//        return (lr - fr + 1) << 1;
//    }
//}
//
//class LongGrid extends FastGrid {
//
//    LongGrid(Dimension dim) {
//        super(dim);
//    }
//
//    @Override
//    int lg() {
//        return lr - fr + 1;
//    }
//}
//
//class IntegerGrid extends FastGrid {
//
//    IntegerGrid(Dimension dim) {
//        super(dim);
//    }
//
//    @Override
//    int lg() {
//        int i = lr - fr + 1;
//        return (i >> 1) + (i & 1);
//    }
//}
//
//class ShortGrid extends FastGrid {
//
//    ShortGrid(Dimension dim) {
//        super(dim);
//    }
//
//    @Override
//    int lg() {
//        int i = lr - fr + 1;
//        return i != (i >> 2 << 2) ? (i >> 2) + 1 : i >> 2;
//    }
//}
//
//class ByteGrid extends FastGrid {
//
//    ByteGrid(Dimension dim) {
//        super(dim);
//    }
//
//    @Override
//    int lg() {
//        int i = lr - fr + 1;
//        return i != (i >> 3 << 3) ? (i >> 3) + 1 : i >> 3;
//    }
//
//}
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
     * @param from the begin index
     * @param to the end index
     */
    void mark(char[] chars, int from, int to);

    /**
     * Mark `1` at the specified {@link Dimension}
     *
     * @param dimension range {@link Dimension}
     */
    void mark(Dimension dimension);
}

class GridFactory {
    private GridFactory() { }
    static Grid create(Dimension dim) {
        // TODO
        return new FractureGrid();
    }
}

abstract class AbstractGrid implements Grid {
    int r, c; // Start index of Row and Column(One base)

}

class FractureGrid implements Grid {

    @Override
    public void mark(char[] chars, int from, int to) {

    }

    @Override
    public void mark(Dimension dimension) {

    }
}

class LongGrid extends AbstractGrid {

    @Override
    public void mark(char[] chars, int from, int to) {

    }

    @Override
    public void mark(Dimension dimension) {

    }
}

class IntegerGrid extends AbstractGrid {

    @Override
    public void mark(char[] chars, int from, int to) {

    }

    @Override
    public void mark(Dimension dimension) {

    }
}

class ShortGrid extends AbstractGrid {

    @Override
    public void mark(char[] chars, int from, int to) {

    }

    @Override
    public void mark(Dimension dimension) {

    }
}

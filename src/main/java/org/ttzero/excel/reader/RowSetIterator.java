/*
 * Copyright (c) 2017-2019, guanquan.wang@yandex.com All Rights Reserved.
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

import java.util.Iterator;
import java.util.NoSuchElementException;
import java.util.function.Supplier;

/**
 * 行迭代器
 *
 * @author guanquan.wang at 2019-04-17 19:03
 */
public class RowSetIterator implements Iterator<Row> {
    protected final Supplier<Row> supplier;
    protected Row nextRow = null;

    public RowSetIterator(Supplier<Row> supplier) {
        this.supplier = supplier;
    }

    @Override
    public boolean hasNext() {
        if (nextRow != null) {
            return true;
        } else {
            nextRow = supplier.get();
            return (nextRow != null);
        }
    }

    @Override
    public Row next() {
        if (nextRow != null || hasNext()) {
            Row next = nextRow;
            nextRow = null;
            return next;
        } else {
            throw new NoSuchElementException();
        }
    }

    public static class NonBlankIterator extends RowSetIterator {

        public NonBlankIterator(Supplier<Row> supplier) {
            super(supplier);
        }

        @Override
        public boolean hasNext() {
            if (nextRow != null) {
                return true;
            } else {
                // Skip blank rows
                for (; (nextRow = supplier.get()) != null && nextRow.isBlank(); ) ;
                return (nextRow != null);
            }
        }
    }
}

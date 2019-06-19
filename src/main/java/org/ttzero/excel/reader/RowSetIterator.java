/*
 * Copyright (c) 2019, guanquan.wang@yandex.com All Rights Reserved.
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

import java.io.IOException;
import java.io.UncheckedIOException;
import java.util.Iterator;
import java.util.NoSuchElementException;

/**
 * Create by guanquan.wang at 2019-04-17 19:03
 */
class RowSetIterator implements Iterator<Row> {
    private boolean onlyDataRow;
    private RowSetProcessor processor;
    private Row nextRow = null;

    public RowSetIterator(RowSetProcessor processor, boolean onlyDataRow) {
        this.processor = processor;
        this.onlyDataRow = onlyDataRow;
    }

    @Override
    public boolean hasNext() {
        if (nextRow != null) {
            return true;
        } else {
            try {
                if (onlyDataRow) {
                    // Skip empty rows
                    for (; (nextRow = processor.next()) != null && nextRow.isEmpty(); ) ;
                } else {
                    nextRow = processor.next();
                }
                return (nextRow != null);
            } catch (IOException e) {
                throw new UncheckedIOException(e);
            }
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
}

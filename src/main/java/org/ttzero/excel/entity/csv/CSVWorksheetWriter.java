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

package org.ttzero.excel.entity.csv;

import org.ttzero.excel.entity.IWorksheetWriter;
import org.ttzero.excel.entity.RowBlock;
import org.ttzero.excel.entity.Sheet;
import org.ttzero.excel.manager.Const;

import java.io.IOException;
import java.nio.file.Path;
import java.util.function.Supplier;

/**
 * Create by guanquan.wang at 2019-08-21 22:19
 */
public class CSVWorksheetWriter implements IWorksheetWriter {
    private Sheet sheet;

    public CSVWorksheetWriter(Sheet sheet) {
        this.sheet = sheet;
    }

    /**
     * The row limit
     *
     * @return the const value {@code (1 << 31) - 1}
     */
    @Override
    public int getRowLimit() {
        return Integer.MAX_VALUE;
    }

    /**
     * The column limit
     *
     * @return the const value 16_384
     */
    @Override
    public int getColumnLimit() {
        return Const.Limit.MAX_COLUMNS_ON_SHEET;
    }

    @Override
    public void writeTo(Path path, Supplier<RowBlock> supplier) throws IOException {

    }

    @Override
    public IWorksheetWriter setWorksheet(Sheet sheet) {
        return null;
    }

    @Override
    public IWorksheetWriter clone() {
        return null;
    }

    @Override
    public void close() throws IOException {

    }

    @Override
    public void writeTo(Path root) throws IOException {

    }
}

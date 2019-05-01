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

package cn.ttzero.excel.entity;

import cn.ttzero.excel.reader.Cell;
import cn.ttzero.excel.util.StringUtil;

import java.io.IOException;
import java.sql.ResultSet;
import java.sql.SQLException;

/**
 * Created by guanquan.wang on 2017/9/27.
 */
public class ResultSetSheet extends Sheet {
    protected ResultSet rs;

    /**
     * Constructor worksheet
     */
    public ResultSetSheet() {
        super();
    }

    /**
     * Constructor worksheet
     * @param name the worksheet name
     */
    public ResultSetSheet(String name) {
        super(name);
    }

    /**
     * Constructor worksheet
     * @param name the worksheet name
     * @param columns the header info
     */
    public ResultSetSheet(String name, final Column[] columns) {
        super(name, columns);
    }

    /**
     * Constructor worksheet
     * @param name the worksheet name
     * @param waterMark the water mark
     * @param columns the header info
     */
    public ResultSetSheet(String name, WaterMark waterMark, final Column[] columns) {
        super(name, waterMark, columns);
    }

    public void setRs(ResultSet rs) {
        this.rs = rs;
    }

    /**
     * Release resources
     * @throws IOException if io error occur
     */
    @Override
    public void close() throws IOException {
        if (shouldClose && rs != null) {
            try {
                rs.close();
            } catch (SQLException e) {
                workbook.what("9006", e.getMessage());
            }
        }
        super.close();
    }

    /**
     * Reset the row-block data
     */
    @Override
    protected void resetBlockData() {
        int len = columns.length, n = 0, limit = sheetWriter.getRowLimit() - 1;

        try {
            for (int rbs = getRowBlockSize(); n++ < rbs && rows < limit && rs.next(); ) {
                Row row = rowBlock.next();
                row.index = rows++;
                Cell[] cells = row.realloc(len);
                for (int i = 1; i <= len; i++) {
                    Column hc = columns[i - 1];

                    // clear cells
                    Cell cell = cells[i - 1];
                    cell.clear();

                    Object e = rs.getObject(i);

                    // blank cell
                    if (e == null) {
                        cell.setBlank();
                        continue;
                    }

                    setCellValueAndStyle(cell, e, hc);
                }
            }
        } catch (SQLException e) {
            throw new ExcelWriteException(e);
        }

        // Paging
        if (rows >= limit) {
            shouldClose = false;
            ResultSetSheet copy = getClass().cast(clone());
            workbook.insertSheet(id, copy);
        } else shouldClose = true;
    }

    /**
     * Returns the header column info
     * @return array of column
     */
    @Override
    public Column[] getHeaderColumns() {
        if (headerReady) return columns;
        if (columns != null) {
            for (int i = 0; i < columns.length; i++) {
                if (StringUtil.isEmpty(columns[i].getName())) {
                    columns[i].setName(String.valueOf(i));
                }
            }
        } else columns = new Column[0];
        headerReady = columns.length > 0;
        return columns;
    }

}

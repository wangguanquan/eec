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

import cn.ttzero.excel.util.StringUtil;

import java.io.IOException;
import java.nio.file.Path;
import java.sql.PreparedStatement;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;

/**
 * Created by guanquan.wang on 2017/9/26.
 */
public class StatementSheet extends ResultSetSheet {
    private PreparedStatement ps;

    /**
     * Constructor worksheet
     */
    public StatementSheet() {
        super();
    }

    /**
     * Constructor worksheet
     * @param name the worksheet name
     */
    public StatementSheet(String name) {
        super(name);
    }

    /**
     * Constructor worksheet
     * @param name the worksheet name
     * @param columns the header info
     */
    public StatementSheet(String name, final Column[] columns) {
        super(name, columns);
    }

    /**
     * Constructor worksheet
     * @param name the worksheet name
     * @param waterMark the water mark
     * @param columns the header info
     */
    public StatementSheet(String name, WaterMark waterMark, final Column[] columns) {
        super(name, waterMark, columns);
    }

    /**
     * @param ps PreparedStatement
     */
    public void setPs(PreparedStatement ps) {
        this.ps = ps;
    }

    /**
     * Release resources
     * @throws IOException if io error occur
     */
    @Override
    public void close() throws IOException {
        super.close();
        if (shouldClose && ps != null) {
            try {
                ps.close();
            } catch (SQLException e) {
                workbook.what("9006", e.getMessage());
            }
        }
    }

    /**
     * write worksheet data to path
     * @param path the storage path
     * @throws IOException write error
     * @throws ExcelWriteException others
     */
    public void writeTo(Path path) throws IOException, ExcelWriteException {
        if (sheetWriter != null) {
            try {
                rs = ps.executeQuery();
            } catch (SQLException e) {
                throw new ExcelWriteException(e);
            }
            rowBlock = new RowBlock();
            sheetWriter.write(path);
        } else {
            throw new ExcelWriteException("Worksheet writer is not instanced.");
        }
    }

    @Override
    public Column[] getHeaderColumns() {
        if (headerReady) return columns;
        // TODO 1.判断各sheet抽出的数据量大小
        int i = 0;
        try {
            ResultSetMetaData metaData = ps.getMetaData();
            for ( ; i < columns.length; i++) {
                if (StringUtil.isEmpty(columns[i].getName())) {
                    columns[i].setName(metaData.getColumnName(i));
                }
                // TODO metaData.getColumnType()
            }
        } catch (SQLException e) {
            workbook.what("un-support get result set meta data.");
        }

        for (i = 0 ; i < columns.length; i++) {
            if (StringUtil.isEmpty(columns[i].getName())) {
                columns[i].setName(String.valueOf(i));
            }
        }
        return columns;
    }
}

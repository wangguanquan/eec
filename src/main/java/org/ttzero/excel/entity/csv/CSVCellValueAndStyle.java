/*
 * Copyright (c) 2017-2019, guanquan.wang@hotmail.com All Rights Reserved.
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

import org.ttzero.excel.entity.Column;
import org.ttzero.excel.entity.ICellValueAndStyle;
import org.ttzero.excel.entity.Row;
import org.ttzero.excel.processor.StyleProcessor;
import org.ttzero.excel.reader.Cell;
import org.ttzero.excel.util.DateUtil;

import java.math.BigDecimal;
import java.sql.Timestamp;
import java.time.format.DateTimeFormatter;

import static org.ttzero.excel.entity.IWorksheetWriter.isBigDecimal;
import static org.ttzero.excel.entity.IWorksheetWriter.isBool;
import static org.ttzero.excel.entity.IWorksheetWriter.isChar;
import static org.ttzero.excel.entity.IWorksheetWriter.isDate;
import static org.ttzero.excel.entity.IWorksheetWriter.isDateTime;
import static org.ttzero.excel.entity.IWorksheetWriter.isDouble;
import static org.ttzero.excel.entity.IWorksheetWriter.isFloat;
import static org.ttzero.excel.entity.IWorksheetWriter.isInt;
import static org.ttzero.excel.entity.IWorksheetWriter.isLocalDate;
import static org.ttzero.excel.entity.IWorksheetWriter.isLocalDateTime;
import static org.ttzero.excel.entity.IWorksheetWriter.isLocalTime;
import static org.ttzero.excel.entity.IWorksheetWriter.isLong;
import static org.ttzero.excel.entity.IWorksheetWriter.isShort;
import static org.ttzero.excel.entity.IWorksheetWriter.isString;
import static org.ttzero.excel.entity.IWorksheetWriter.isTime;
import static org.ttzero.excel.util.DateUtil.toDateString;
import static org.ttzero.excel.util.DateUtil.toDateTimeString;

/**
 * CSV数据样式转换器，该转换器将所有数据转为字符器格式并忽略所有样式
 *
 * @author guanquan.wang at 2019-09-25 11:46
 */
public class CSVCellValueAndStyle implements ICellValueAndStyle {

    /**
     * Setting all cell style of the specified row
     *
     * @param <T> the row's class
     * @param o the row data
     * @param cell the cell of row
     * @param hc the header column
     * @param styleProcessor a customize {@link StyleProcessor}
     */
    @Override
    public <T> void setStyleDesign(T o, Cell cell, Column hc, StyleProcessor<T> styleProcessor) { }

    /**
     * Setting cell value
     *
     * @param cell  the cell
     * @param e     the cell value
     * @param hc    the header column
     * @param clazz the cell value type
     * @param hasConversion 是否有输出转换器
     */
    @Override
    public void setCellValue(Row row, Cell cell, Object e, Column hc, Class<?> clazz, boolean hasConversion) {
        if (hasConversion) {
            conversion(row, cell, e, hc);
            return;
        }
        if (e == null) {
            setNullValue(row, cell, hc);
            return;
        }
        if (isString(clazz)) {
            cell.setString(e.toString());
        } else if (isDate(clazz)) {
            // TODO hc.numFmt
            cell.setString(toDateString((java.util.Date) e));
        } else if (isDateTime(clazz)) {
            cell.setString(toDateTimeString((Timestamp) e));
        } else if (isChar(clazz)) {
            cell.setChar((Character) e);
        } else if (isShort(clazz)) {
            cell.setInt((Short) e);
        } else if (isInt(clazz)) {
            cell.setInt((Integer) e);
        } else if (isLong(clazz)) {
            cell.setLong((Long) e);
        } else if (isFloat(clazz)) {
            cell.setDouble((Float) e);
        } else if (isDouble(clazz)) {
            cell.setDouble((Double) e);
        } else if (isBool(clazz)) {
            cell.setBool((Boolean) e);
        } else if (isBigDecimal(clazz)) {
            cell.setDecimal((BigDecimal) e);
        } else if (isLocalDate(clazz)) {
            cell.setString(((java.time.LocalDate) e).toString());
        } else if (isLocalDateTime(clazz)) {
            cell.setString(DateUtil.LOCAL_DATE_TIME.format((java.time.LocalDateTime) e));
        } else if (isTime(clazz)) {
            cell.setString(DateTimeFormatter.ISO_TIME.format(((java.sql.Time) e).toLocalTime()));
        } else if (isLocalTime(clazz)) {
            cell.setString(DateTimeFormatter.ISO_TIME.format((java.time.LocalTime) e));
        } else {
            cell.setString(e.toString());
        }
    }

    /**
     * Returns the cell style index
     *
     * @param row the row data
     * @param hc the header column
     * @param o  the cell value
     * @return const zero (general style)
     */
    @Override
    public int getStyleIndex(Row row, Column hc, Object o) {
        return 0;
    }

}

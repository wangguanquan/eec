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

package org.ttzero.excel.entity.csv;

import org.ttzero.excel.entity.Column;
import org.ttzero.excel.entity.ICellValueAndStyle;
import org.ttzero.excel.entity.Sheet;
import org.ttzero.excel.manager.Const;
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

/**
 * @author guanquan.wang at 2019-09-25 11:46
 */
public class CSVCellValueAndStyle implements ICellValueAndStyle {
    /**
     * Setting cell value and cell styles
     *
     * @param cell the cell
     * @param e    the cell value
     * @param hc   the header column
     */
    @Override
    public void reset(int row, Cell cell, Object e, Column hc) {
        setCellValue(row, cell, e, hc, hc.getClazz(), hc.processor != null);
    }

    @Override
    public <T> void setStyleDesign(T o, Cell cell, Column hc, Sheet sheet) {

    }

    /**
     * Returns the worksheet name
     *
     * @return name of worksheet
     */
    @Override
    public String getFileSuffix() {
        return Const.Suffix.CSV;
    }

    /**
     * Setting cell value
     *
     * @param cell  the cell
     * @param e     the cell value
     * @param hc    the header column
     * @param clazz the cell value type
     */
    @Override
    public void setCellValue(int row, Cell cell, Object e, Column hc, Class<?> clazz, boolean hasProcessor) {
        if (hasProcessor) {
            conversion(row, cell, e, hc);
            return;
        }
        if (e == null) {
            setNullValue(row, cell, hc);
            return;
        }
        if (isString(clazz)) {
            cell.setSv(e.toString());
        } else if (isDate(clazz)) {
            // TODO hc.numFmt
            cell.setSv(DateUtil.toDateString((java.util.Date) e));
        } else if (isDateTime(clazz)) {
            cell.setSv(DateUtil.toString((Timestamp) e));
        } else if (isChar(clazz)) {
            cell.setCv((Character) e);
        } else if (isShort(clazz)) {
            cell.setNv((Short) e);
        } else if (isInt(clazz)) {
            cell.setNv((Integer) e);
        } else if (isLong(clazz)) {
            cell.setLv((Long) e);
        } else if (isFloat(clazz)) {
            cell.setDv((Float) e);
        } else if (isDouble(clazz)) {
            cell.setDv((Double) e);
        } else if (isBool(clazz)) {
            cell.setBv((Boolean) e);
        } else if (isBigDecimal(clazz)) {
            cell.setMv((BigDecimal) e);
        } else if (isLocalDate(clazz)) {
            cell.setSv(((java.time.LocalDate) e).toString());
        } else if (isLocalDateTime(clazz)) {
            cell.setSv(DateUtil.LOCAL_DATE_TIME.format((java.time.LocalDateTime) e));
        } else if (isTime(clazz)) {
            cell.setSv(DateTimeFormatter.ISO_TIME.format(((java.sql.Time) e).toLocalTime()));
        } else if (isLocalTime(clazz)) {
            cell.setSv(DateTimeFormatter.ISO_TIME.format((java.time.LocalTime) e));
        } else {
            cell.setSv(e.toString());
        }
    }

    /**
     * Returns the cell style index
     *
     * @param hc the header column
     * @param o  the cell value
     * @return const zero (general style)
     */
    @Override
    public int getStyleIndex(int rows, Column hc, Object o) {
        return 0;
    }
}

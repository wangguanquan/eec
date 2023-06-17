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

package org.ttzero.excel.entity;

import org.ttzero.excel.processor.StyleProcessor;
import org.ttzero.excel.reader.Cell;
import org.ttzero.excel.util.DateUtil;
import org.ttzero.excel.util.StringUtil;

import java.io.File;
import java.io.InputStream;
import java.math.BigDecimal;
import java.nio.file.Path;
import java.sql.Timestamp;
import java.util.Base64;

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
 * @author guanquan.wang at 2019-09-25 11:24
 */
public interface ICellValueAndStyle {
    /**
     * Setting cell value and cell styles
     *
     * @param row  the row number
     * @param cell the cell
     * @param e    the cell value
     * @param hc   the header column
     */
    default void reset(Row row, Cell cell, Object e, Column hc) {
        setCellValue(row.index, cell, e, hc, hc.getClazz(), hc.processor != null);
        // Cell style
        if (hc.processor == null) {
            cell.xf = getStyleIndex(row, hc, e);
        }
        // Reset row height
    }

    /**
     * Returns the cell style index
     *
     * @param row the row data
     * @param hc    the header column
     * @param o     the cell value
     * @return the style index in xf
     */
    default int getStyleIndex(Row row, Column hc, Object o) {
        int style = hc.getCellStyle();
        if (hc.styleProcessor != null) {
            style = hc.styleProcessor.build(o, style, hc.styles);
        }
        return hc.styles.of(style);
    }

    /**
     * Setting cell value and cell styles
     *
     * @param row  the row number
     * @param cell the cell
     * @param e    the cell value
     * @param hc   the header column
     * @deprecated Replace with {@link #reset(Row, Cell, Object, Column)}
     */
    @Deprecated
    void reset(int row, Cell cell, Object e, Column hc);

    /**
     * Returns the worksheet name
     *
     * @return name of worksheet
     */
    String getFileSuffix();

    /**
     * Returns the cell style index
     *
     * @param rows the row number
     * @param hc the header column
     * @param o  the cell value
     * @return the style index in xf
     * @deprecated Replace with {@link #getStyleIndex(Row, Column, Object)}
     */
    @Deprecated
    int getStyleIndex(int rows, Column hc, Object o);

    /**
     * Setting all cell style of the specified row
     *
     * @param <T> the row's class
     * @param o the row data
     * @param cell the cell of row
     * @param hc the header column
     * @param styleProcessor a customize {@link StyleProcessor}
     */
    default <T> void setStyleDesign(T o, Cell cell, Column hc, StyleProcessor<T> styleProcessor) {
        if (styleProcessor != null && hc.styles != null) {
            cell.xf = hc.styles.of(styleProcessor.build(o, hc.styles.getStyleByIndex(cell.xf), hc.styles));
        }
    }

    /**
     * Setting cell value
     *
     * @param row the row number
     * @param cell  the cell
     * @param e     the cell value
     * @param hc    the header column
     * @param clazz the cell value type
     * @param hasProcessor Specify the cell has value converter
     */
    default void setCellValue(int row, Cell cell, Object e, Column hc, Class<?> clazz, boolean hasProcessor) {
        if (hasProcessor) {
            conversion(row, cell, e, hc);
            return;
        }
        if (e == null) {
            setNullValue(row, cell, hc);
            return;
        }
        if (clazz == null) {
            clazz = e.getClass();
            hc.setClazz(clazz);
        }
        if (isString(clazz)) {
            // If write as base64 images
            // base64Image(row, cell, e.toString(), hc, clazz);
            cell.setSv(e.toString());
        } else if (isDate(clazz)) {
            cell.setIv(DateUtil.toDateTimeValue((java.util.Date) e));
        } else if (isDateTime(clazz)) {
            cell.setIv(DateUtil.toDateTimeValue((Timestamp) e));
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
            cell.setIv(DateUtil.toDateValue((java.time.LocalDate) e));
        } else if (isLocalDateTime(clazz)) {
            cell.setIv(DateUtil.toDateTimeValue((java.time.LocalDateTime) e));
        } else if (isTime(clazz)) {
            cell.setTv(DateUtil.toTimeValue((java.sql.Time) e));
        } else if (isLocalTime(clazz)) {
            cell.setTv(DateUtil.toTimeValue((java.time.LocalTime) e));
        }
        // TODO check column export type is image
        else if (clazz == byte[].class) {
            cell.setBinary((byte[]) e);
        } else if (clazz == Path.class) {
            cell.setPath((Path) e);
        } else if (clazz == File.class) {
            cell.setPath(((File) e).toPath());
        } else if (InputStream.class.isAssignableFrom(clazz)) {
            cell.setInputStream((InputStream) e);
        } else {
            unknownType(row, cell, e, hc, clazz);
        }
    }

    /**
     * Setting cell value as null
     *
     * @param row the row number
     * @param cell  the cell
     * @param hc    the header column
     */
    default void setNullValue(int row, Cell cell, Column hc) {
        boolean hasProcessor = hc.processor != null;
        if (hasProcessor) {
            conversion(row, cell, 0, hc);
        } else
            cell.blank();
    }

    /**
     * Int value conversion to others
     *
     * @param row the row number
     * @param cell the cell
     * @param o    the cell value
     * @param hc   the header column
     */
    default void conversion(int row, Cell cell, Object o, Column hc) {
        Object e = hc.processor.conversion(o);
        if (e != null) {
            Class<?> clazz = e.getClass();
            if (isInt(clazz)) {
                if (isChar(clazz)) {
                    cell.setCv((Character) e);
                } else if (isShort(clazz)) {
                    cell.setNv((Short) e);
                } else {
                    cell.setNv((Integer) e);
                }
            } else {
                setCellValue(row, cell, e, hc, clazz, false);
            }
        } else {
            cell.blank();
        }
    }

    /**
     * unknown cell type converter
     *
     * @param row the row number
     * @param cell  the cell
     * @param e     the cell value
     * @param hc    the header column
     * @param clazz the cell value type
     */
    default void unknownType(int row, Cell cell, Object e, Column hc, Class<?> clazz) {
        cell.setSv(e.toString());
    }

    /**
     * Base64 string convert to binary
     *
     * @param row the row number
     * @param cell  the cell
     * @param e     the cell value
     * @param hc    the header column
     * @param clazz the cell value type
     */
    default void base64Image(int row, Cell cell, String e, Column hc, Class<?> clazz) {
        int b;
        if (e.startsWith("data:") && (b = StringUtil.indexOf(e, ',', 6, 64)) > 6
            && (e.charAt(b - 1) == '4' && e.charAt(b - 2) == '6' && e.charAt(b - 3) == 'e' && e.charAt(b - 4) == 's' && e.charAt(b - 5) == 'a' && e.charAt(b - 6) == 'b')) {
            byte[] bytes = Base64.getDecoder().decode(e.substring(b + 1));
            cell.setBinary(bytes);
        } else cell.setSv(e);
    }
}

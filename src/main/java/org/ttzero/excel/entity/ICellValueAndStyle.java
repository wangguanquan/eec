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
import java.nio.ByteBuffer;
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
        boolean hasConversion = hc.getConversion() != null;
        setCellValue(row.index, cell, e, hc, hc.getClazz(), hasConversion);
        // Cell style
        if (!hasConversion) {
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
     * @param hasConversion 是否有输出转换器
     */
    default void setCellValue(int row, Cell cell, Object e, Column hc, Class<?> clazz, boolean hasConversion) {
        if (hasConversion) {
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
            switch (hc.getColumnType()) {
                // Default
                case 0: cell.setSv(e.toString()); break;
                // Write as media (base64 image, remote url)
                case 1: writeAsMedia(row, cell, e.toString(), hc, clazz); break;
                default: cell.setSv(e.toString());
            }
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
        // Write as media if column-type equals {@code 1}
        else if (hc.getColumnType() == 1) {
            if (Path.class.isAssignableFrom(clazz)) {
                cell.setPath((Path) e);
            } else if (File.class.isAssignableFrom(clazz)) {
                 cell.setPath(((File) e).toPath());
            } else if (InputStream.class.isAssignableFrom(clazz)) {
                cell.setInputStream((InputStream) e);
            } else if (clazz == byte[].class) {
                cell.setBinary((byte[]) e);
            } else if (ByteBuffer.class.isAssignableFrom(clazz)) {
                cell.setByteBuffer((ByteBuffer) e);
            }
        }
        // Others
        else {
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
        boolean hasProcessor = hc.getConversion() != null;
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
        Object e = hc.getConversion().conversion(o);
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
     * Convert string to binary
     *
     * @param row   the row number
     * @param cell  the cell
     * @param e     the cell value
     * @param hc    the header column
     * @param clazz the cell value type
     */
    default void writeAsMedia(int row, Cell cell, String e, Column hc, Class<?> clazz) {
        int b, len = e.length();
        // Base64 image
        if (len > 64 && e.startsWith("data:") && (b = StringUtil.indexOf(e, ',', 6, 64)) > 6
            && (e.charAt(b - 1) == '4' && e.charAt(b - 2) == '6' && e.charAt(b - 3) == 'e' && e.charAt(b - 4) == 's' && e.charAt(b - 5) == 'a' && e.charAt(b - 6) == 'b')) {
            byte[] bytes = Base64.getDecoder().decode(e.substring(b + 1));
            cell.setBinary(bytes);
        }
        // Remote uri (http:// | https:// | ftp:// | ftps://)
        else if (len >= 10 && (b = StringUtil.indexOf(e, ':', 3, 6)) >= 3 && e.charAt(b + 1) == '/' && e.charAt(b + 2) == '/') {
            downloadRemoteResource(row, cell, e, hc, clazz);
        }
        // Others
        else cell.setSv(e);
    }

    /**
     * Download resource from remote server
     *
     * NOTE：By default, only marking the cell type as {@code REMOTE_URL} does not actually download resources.
     * Asynchronous download in {@code IWorksheetWriter#writeRemoteMedia} in the future.
     * Of course, you can also download and then call {@link Cell#setInputStream(InputStream)}
     * or {@link Cell#setBinary(byte[])} to save the stream or binary results to the cell.
     *
     * @param row   the row number
     * @param cell  the cell
     * @param e     the cell value
     * @param hc    the header column
     * @param clazz the cell value type
     */
    default void downloadRemoteResource(int row, Cell cell, String e, Column hc, Class<?> clazz) {
        cell.setSv(e);
        cell.t = Cell.REMOTE_URL;
    }

    /**
     * Mark whitelist types that can be easily exported
     *
     * @param clazz cell value class
     * @return true if can be easily exported
     */
    default boolean isAllowDirectOutput(Class<?> clazz) {
        return clazz == null || isString(clazz) || isDate(clazz) || isDateTime(clazz) || isChar(clazz) || isShort(clazz)
            || isInt(clazz) || isLong(clazz) || isFloat(clazz) || isDouble(clazz) || isBool(clazz) || isBigDecimal(clazz)
            || isLocalDate(clazz) || isLocalDateTime(clazz) || isTime(clazz) || isLocalTime(clazz) || Path.class.isAssignableFrom(clazz)
            || File.class.isAssignableFrom(clazz) || InputStream.class.isAssignableFrom(clazz) || clazz == byte[].class
            || ByteBuffer.class.isAssignableFrom(clazz);
    }
}

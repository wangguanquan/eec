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
 * 数据转换并设置样式，将数据写到工作表输出协议之前统一处理数据和样式
 *
 * @author guanquan.wang at 2019-09-25 11:24
 */
public interface ICellValueAndStyle {
    /**
     * 重置单元格的值和样式，Row和Cell都是内存共享的，所以每个单元格均需要重置值和样式
     *
     * @param row  行信息
     * @param cell 单元格
     * @param e    单元格的值
     * @param hc   当前列的表头
     */
    default void reset(Row row, Cell cell, Object e, Column hc) {
        // 将值转输出需要的统一格式
        setCellValue(row, cell, e, hc, hc.getClazz(), hc.getConversion() != null);
        // 单元格样式
        cell.xf = getStyleIndex(row, hc, e);
    }

    /**
     * 获取单元格样式值，先通过{@code Column}获取基础样式，如果有动态样式转换则将基础样式做为参数进行二次制作
     *
     * @param row 行信息
     * @param hc  当前列的表头
     * @param o   单元格的值
     * @return 样式值
     */
    default int getStyleIndex(Row row, Column hc, Object o) {
        // 获取基础样式
        int style = hc.getCellStyle();
        // 如果有动态样式转换则将基础样式做为参数进行二次制作
        if (hc.styleProcessor != null) {
            style = hc.styleProcessor.build(o, style, hc.styles);
        }
        return hc.styles.of(style);
    }

    /**
     * 行级样式转换器，它的优先级最高
     *
     * @param <T>            the row's class
     * @param o              行值，可能是{@code Map}，Java实体或{@code ResultSet}
     * @param cell           单元格
     * @param hc             当前列的表头
     * @param styleProcessor 样式转换器{@link StyleProcessor}
     */
    default <T> void setStyleDesign(T o, Cell cell, Column hc, StyleProcessor<T> styleProcessor) {
        if (styleProcessor != null && hc.styles != null) {
            cell.xf = hc.styles.of(styleProcessor.build(o, hc.styles.getStyleByIndex(cell.xf), hc.styles));
        }
    }

    /**
     * 设置单元格的值，如果有动态转换器则调用转换器
     *
     * @param row           行信息
     * @param cell          单元格
     * @param e             单元格的值
     * @param hc            当前列的表头
     * @param clazz         单元格值的数据类型
     * @param hasConversion 是否有输出转换器
     */
    default void setCellValue(Row row, Cell cell, Object e, Column hc, Class<?> clazz, boolean hasConversion) {
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
                case 0: cell.setString(e.toString()); break;
                // Write as media (base64 image, remote url)
                case 1: writeAsMedia(row, cell, e.toString(), hc, clazz); break;
                default: cell.setString(e.toString());
            }
        } else if (isDate(clazz)) {
            cell.setDateTime(DateUtil.toDateTimeValue((java.util.Date) e));
        } else if (isDateTime(clazz)) {
            cell.setDateTime(DateUtil.toDateTimeValue((Timestamp) e));
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
            cell.setDateTime(DateUtil.toDateValue((java.time.LocalDate) e));
        } else if (isLocalDateTime(clazz)) {
            cell.setDateTime(DateUtil.toDateTimeValue((java.time.LocalDateTime) e));
        } else if (isTime(clazz)) {
            cell.setTime(DateUtil.toTimeValue((java.sql.Time) e));
        } else if (isLocalTime(clazz)) {
            cell.setTime(DateUtil.toTimeValue((java.time.LocalTime) e));
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
     * 写{@code null}值到单元格
     *
     * @param row  行信息
     * @param cell 单元格
     * @param hc   当前列的表头
     */
    default void setNullValue(Row row, Cell cell, Column hc) {
        boolean hasProcessor = hc.getConversion() != null;
        if (hasProcessor) {
            conversion(row, cell, 0, hc);
        } else
            cell.blank();
    }

    /**
     * 动态转换单元格的值
     *
     * @param row  行信息
     * @param cell 单元格
     * @param o    单元格的值
     * @param hc   当前列的表头
     */
    default void conversion(Row row, Cell cell, Object o, Column hc) {
        Object e = hc.getConversion().conversion(o);
        if (e != null) {
            setCellValue(row, cell, e, hc, e.getClass(), false);
        } else {
            cell.blank();
        }
    }

    /**
     * 未知类型转换，可覆写本方法以支持扩展类型
     *
     * @param row   行信息
     * @param cell  单元格
     * @param e     单元格的值
     * @param hc    当前列的表头
     * @param clazz 单元格值的数据类型
     */
    default void unknownType(Row row, Cell cell, Object e, Column hc, Class<?> clazz) {
        cell.setString(e.toString());
    }

    /**
     * 将字符串转为{@code Media}类型，仅当以{@code Media}类型导出时才会被执行，
     * 支持Base64图片和图片url，
     *
     * @param row   行信息
     * @param cell  单元格
     * @param e     单元格的值
     * @param hc    当前列的表头
     * @param clazz 单元格值的数据类型
     */
    default void writeAsMedia(Row row, Cell cell, String e, Column hc, Class<?> clazz) {
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
        else cell.setString(e);
    }

    /**
     * 下载远程资源
     *
     * <p>注意：默认情况下仅将单元格类型标记为{@code REMOTE_URL}并不会去下载资源。
     * 下载动作延迟在{@code IWorksheetWriter#writeRemoteMedia}中进行。
     * 当然，也可以在本方法下载并调用{@link Cell#setInputStream}或{@link Cell#setBinary}
     * 将流或二进制结果保存到单元格中</p>
     *
     * @param row   行信息
     * @param cell  单元格
     * @param e     单元格的值
     * @param hc    当前列的表头
     * @param clazz 单元格值的数据类型
     */
    default void downloadRemoteResource(Row row, Cell cell, String e, Column hc, Class<?> clazz) {
        cell.setString(e);
        cell.t = Cell.REMOTE_URL;
    }

    /**
     * 检查数据类型是否可简单导出，简单导出的类型是相对于实体而言，它们一定是Java内置类型且被其它实体组合使用
     *
     * @param clazz 数据类型
     * @return {@code true}如果是简单类型
     */
    default boolean isAllowDirectOutput(Class<?> clazz) {
        return clazz == null || isString(clazz) || isDate(clazz) || isDateTime(clazz) || isChar(clazz) || isShort(clazz)
            || isInt(clazz) || isLong(clazz) || isFloat(clazz) || isDouble(clazz) || isBool(clazz) || isBigDecimal(clazz)
            || isLocalDate(clazz) || isLocalDateTime(clazz) || isTime(clazz) || isLocalTime(clazz) || Path.class.isAssignableFrom(clazz)
            || File.class.isAssignableFrom(clazz) || InputStream.class.isAssignableFrom(clazz) || clazz == byte[].class
            || ByteBuffer.class.isAssignableFrom(clazz);
    }
}

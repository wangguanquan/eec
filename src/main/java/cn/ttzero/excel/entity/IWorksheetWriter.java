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

import java.io.Closeable;
import java.io.IOException;
import java.math.BigDecimal;
import java.nio.file.Path;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.function.Supplier;

/**
 * Create by guanquan.wang at 2019-04-22 17:23
 */
public interface IWorksheetWriter extends Closeable {

    /**
     * The Worksheet row limit
     * @return the limit
     */
    int getRowLimit();

    /**
     * The Worksheet column limit
     * @return the limit
     */
    int getColumnLimit();

    /**
     * Write a row block
     * @param path the storage path
     * @param supplier a row-block supplier
     * @throws IOException if io error occur
     */
    void write(Path path, Supplier<RowBlock> supplier) throws IOException;

    /**
     * Write a row block
     * @param path the storage path
     * @throws IOException if io error occur
     */
    void write(Path path) throws IOException;

    /**
     * Write a empty worksheet
     * @param path the path to storage
     * @throws IOException if io error occur
     */
    default void writeEmptySheet(Path path) throws IOException {
        try {
            write(path, () -> null);
        } finally {
            close();
        }
    }

    /**
     * TO check rows out of workshee
     * @param row row number
     * @return true if rows large than limit
     */
    default boolean outOfSheet(int row) {
        return row >= getRowLimit();
    }

    /**
     * 测试是否为Date类型
     * @param clazz 列类型
     * @return bool
     */
    static boolean isDate(Class<?> clazz) {
        return clazz == java.util.Date.class
            || clazz == java.sql.Date.class;
    }

    /**
     * 测试是否为DateTime类型
     * @param clazz 列类型
     * @return bool
     */
    static boolean isDateTime(Class<?> clazz) {
        return clazz == java.sql.Timestamp.class;
    }

    /**
     * 测试是否为Int类型
     * @param clazz 列类型
     * @return bool
     */
    static boolean isInt(Class<?> clazz) {
        return clazz == int.class || clazz == Integer.class
            || clazz == char.class || clazz == Character.class
            || clazz == byte.class || clazz == Byte.class
            || clazz == short.class || clazz == Short.class;
    }

    /**
     * Test clazz is short class
     * @param clazz the class
     * @return boolean value
     */
    static boolean isShort(Class<?> clazz) {
        return clazz == short.class || clazz == Short.class;
    }

    /**
     * 测试是否为Long类型
     * @param clazz 列类型
     * @return bool
     */
    static boolean isLong(Class<?> clazz) {
        return clazz == long.class || clazz == Long.class;
    }

    /**
     * 测试是否为Float类型
     * @param clazz 列类型
     * @return bool
     */
    static boolean isFloat(Class<?> clazz) {
        return clazz == float.class || clazz == Float.class;
    }

    /**
     * Test clazz is double class
     * @param clazz the class
     * @return boolean value
     */
    static boolean isDouble(Class<?> clazz) {
        return clazz == double.class || clazz == Double.class;
    }

    /**
     * 测试是否为Boolean类型
     * @param clazz 列类型
     * @return bool
     */
    static boolean isBool(Class<?> clazz) {
        return clazz == boolean.class || clazz == Boolean.class;
    }

    /**
     * 测试是否为String类型
     * @param clazz 列类型
     * @return bool
     */
    static boolean isString(Class<?> clazz) {
        return clazz == String.class || clazz == CharSequence.class;
    }

    /**
     * 测试是否为Char类型
     * @param clazz 列类型
     * @return bool
     */
    static boolean isChar(Class<?> clazz) {
        return clazz == char.class || clazz == Character.class;
    }

    /**
     * 测试是否为BigDecimal类型
     * @param clazz 列类型
     * @return bool
     */
    static boolean isBigDecimal(Class<?> clazz) {
        return clazz == BigDecimal.class;
    }

    /**
     * 测试是否为LocalDate类型
     * @param clazz 列类型
     * @return bool
     */
    static boolean isLocalDate(Class<?> clazz) {
        return clazz == LocalDate.class;
    }

    /**
     * 测试是否为LocalDateTime类型
     * @param clazz 列类型
     * @return bool
     */
    static boolean isLocalDateTime(Class<?> clazz) {
        return clazz == LocalDateTime.class;
    }

    /**
     * 测试是否为java.sql.Time类型
     * @param clazz 列类型
     * @return bool
     */
    static boolean isTime(Class<?> clazz) {
        return clazz == java.sql.Time.class;
    }

    /**
     * 测试是否为LocalTime类型
     * @param clazz 列类型
     * @return bool
     */
    static boolean isLocalTime(Class<?> clazz) {
        return clazz == java.time.LocalTime.class;
    }

}

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

import java.io.Closeable;
import java.io.IOException;
import java.math.BigDecimal;
import java.nio.file.Path;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.function.Supplier;

/**
 * @author guanquan.wang at 2019-04-22 17:23
 */
public interface IWorksheetWriter extends Closeable, Cloneable, Storable {

    /**
     * The Worksheet row limit
     *
     * @return the limit
     */
    int getRowLimit();

    /**
     * The Worksheet column limit
     *
     * @return the limit
     */
    int getColumnLimit();

    /**
     * Write a row block
     *
     * @param path     the storage path
     * @param supplier a row-block supplier
     * @throws IOException if io error occur
     */
    void writeTo(Path path, Supplier<RowBlock> supplier) throws IOException;

    /**
     * Return a copy worksheet writer
     *
     * @param sheet the {@link Sheet}
     * @return the copy worksheet writer
     */
    IWorksheetWriter setWorksheet(Sheet sheet);

    /**
     * Write a empty worksheet
     *
     * @param path the path to storage
     * @throws IOException if io error occur
     */
    default void writeEmptySheet(Path path) throws IOException {
        try {
            writeTo(path, () -> null);
        } finally {
            close();
        }
    }

    /**
     * TO check rows out of worksheet
     *
     * @param row row number
     * @return true if rows large than limit
     */
    default boolean isOutOfSheet(int row) {
        return row >= getRowLimit();
    }

    /**
     * Clone
     *
     * @return IWorksheetWriter
     */
    IWorksheetWriter clone();

    /**
     * Returns the worksheet name
     *
     * @return name of worksheet
     */
    String getFileSuffix();

    /**
     * Test if it is a {@link java.util.Date} type
     *
     * @param clazz the type
     * @return bool
     */
    static boolean isDate(Class<?> clazz) {
        return clazz == java.util.Date.class
            || clazz == java.sql.Date.class;
    }

    /**
     * Test if it is a {@link java.sql.Timestamp} type
     *
     * @param clazz the type
     * @return bool
     */
    static boolean isDateTime(Class<?> clazz) {
        return clazz == java.sql.Timestamp.class;
    }

    /**
     * Test if it is a {@code int, char, byte or short} or boxing type
     *
     * @param clazz the type
     * @return bool
     */
    static boolean isInt(Class<?> clazz) {
        return clazz == int.class || clazz == Integer.class
            || clazz == char.class || clazz == Character.class
            || clazz == byte.class || clazz == Byte.class
            || clazz == short.class || clazz == Short.class;
    }

    /**
     * Test if it is a {@code short} or {@link Short} type
     *
     * @param clazz the class
     * @return boolean value
     */
    static boolean isShort(Class<?> clazz) {
        return clazz == short.class || clazz == Short.class;
    }

    /**
     * Test if it is a {@code long} or {@link Long} type
     *
     * @param clazz the type
     * @return bool
     */
    static boolean isLong(Class<?> clazz) {
        return clazz == long.class || clazz == Long.class;
    }

    /**
     * Test if it is a single-precision floating-point type
     *
     * @param clazz the type
     * @return bool
     */
    static boolean isFloat(Class<?> clazz) {
        return clazz == float.class || clazz == Float.class;
    }

    /**
     * Test if it is a double-precision floating-point type
     *
     * @param clazz the type
     * @return boolean value
     */
    static boolean isDouble(Class<?> clazz) {
        return clazz == double.class || clazz == Double.class;
    }

    /**
     * Test if it is a {@code boolean} or {@link Boolean} type
     *
     * @param clazz the type
     * @return bool
     */
    static boolean isBool(Class<?> clazz) {
        return clazz == boolean.class || clazz == Boolean.class;
    }

    /**
     * Test if it is a {@link String} type
     *
     * @param clazz the type
     * @return bool
     */
    static boolean isString(Class<?> clazz) {
        return clazz == String.class || clazz == CharSequence.class;
    }

    /**
     * Test if it is a {@code char} or {@link Character} type
     *
     * @param clazz the type
     * @return bool
     */
    static boolean isChar(Class<?> clazz) {
        return clazz == char.class || clazz == Character.class;
    }

    /**
     * Test if it is a {@link BigDecimal} type
     *
     * @param clazz the type
     * @return bool
     */
    static boolean isBigDecimal(Class<?> clazz) {
        return clazz == BigDecimal.class;
    }

    /**
     * Test if it is a {@link LocalDate} type
     *
     * @param clazz the type
     * @return bool
     */
    static boolean isLocalDate(Class<?> clazz) {
        return clazz == LocalDate.class;
    }

    /**
     * Test if it is a {@link LocalDateTime} type
     *
     * @param clazz the type
     * @return bool
     */
    static boolean isLocalDateTime(Class<?> clazz) {
        return clazz == LocalDateTime.class;
    }

    /**
     * Test if it is a {@link java.sql.Time} type
     *
     * @param clazz the type
     * @return bool
     */
    static boolean isTime(Class<?> clazz) {
        return clazz == java.sql.Time.class;
    }

    /**
     * Test if it is a {@link java.time.LocalTime} type
     *
     * @param clazz the type
     * @return bool
     */
    static boolean isLocalTime(Class<?> clazz) {
        return clazz == java.time.LocalTime.class;
    }

}

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

package cn.ttzero.excel.reader;

import java.math.BigDecimal;
import java.sql.Timestamp;
import java.util.Date;

/**
 * Create by guanquan.wang at 2019-04-17 11:08
 */
public interface Row {
    /**
     * The number of row. (zero base)
     * @return int value
     */
    int getRowNumber();

    /**
     * Test unused row (not contains any filled or formatted or value)
     * @return true if unused
     */
    boolean isEmpty();

    /**
     * Get boolean value by column index
     * @param columnIndex the cell index
     * @return boolean
     */
    boolean getBoolean(int columnIndex);

    /**
     * Get byte value by column index
     * @param columnIndex the cell index
     * @return byte
     */
    byte getByte(int columnIndex);

    /**
     * Get char value by column index
     * @param columnIndex the cell index
     * @return char
     */
    char getChar(int columnIndex);

    /**
     * Get short value by column index
     * @param columnIndex the cell index
     * @return short
     */
    short getShort(int columnIndex);

    /**
     * Get int value by column index
     * @param columnIndex the cell index
     * @return int
     */
    int getInt(int columnIndex);

    /**
     * Get long value by column index
     * @param columnIndex the cell index
     * @return long
     */
    long getLong(int columnIndex);

    /**
     * Get string value by column index
     * @param columnIndex the cell index
     * @return string
     */
    String getString(int columnIndex);

    /**
     * Get float value by column index
     * @param columnIndex the cell index
     * @return float
     */
    float getFloat(int columnIndex);

    /**
     * Get double value by column index
     * @param columnIndex the cell index
     * @return double
     */
    double getDouble(int columnIndex);

    /**
     * Get decimal value by column index
     * @param columnIndex the cell index
     * @return BigDecimal
     */
    BigDecimal getDecimal(int columnIndex);

    /**
     * Get date value by column index
     * @param columnIndex the cell index
     * @return Date
     */
    Date getDate(int columnIndex);

    /**
     * Get timestamp value by column index
     * @param columnIndex the cell index
     * @return java.sql.Timestamp
     */
    Timestamp getTimestamp(int columnIndex);

    /**
     * Get time value by column index
     * @param columnIndex the cell index
     * @return java.sql.Time
     */
    java.sql.Time getTime(int columnIndex);

    /**
     * Get T value by column index
     * Override this method
     * @param columnIndex the cell index
     * @param <T> the type of return object
     * @return T
     */
    default <T> T get(int columnIndex) {
        throw new UnsupportedOperationException();
    }

    /**
     * Returns the binding type if is bound, otherwise returns Row
     * @param <T> the type of binding
     * @return T
     */
    <T> T get();

    /**
     * Returns the binding type if is bound, otherwise returns Row
     * @param <T> the type of binding
     * @return T
     */
    <T> T geet();

    /**
     * Convert to object, support annotation
     * @param clazz the type of binding
     * @param <T> the type of return object
     * @return T
     */
    <T> T to(Class<T> clazz);

    /**
     * Convert to T object, support annotation
     * the is a memory shared object
     * @param clazz the type of binding
     * @param <T> the type of return object
     * @return T
     */
    <T> T too(Class<T> clazz);
}

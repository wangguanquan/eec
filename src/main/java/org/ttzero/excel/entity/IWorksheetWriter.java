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

package org.ttzero.excel.entity;

import org.ttzero.excel.entity.e7.XMLCellValueAndStyle;

import java.io.Closeable;
import java.io.IOException;
import java.math.BigDecimal;
import java.nio.file.Path;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.function.Supplier;

/**
 * 工作表输出协议，负责将工作表{@link Sheet}格式化输出，它会循环调用{@link Sheet#nextBlock}
 * 方法获取数据并写入磁盘直到{@link RowBlock#isEOF}返回EOF标记为止，整个过程只有一个
 * RowBlock行块常驻内存，一个{@code RowBlock}行块默认包含32个{@code Row}行，这样可以保证
 * 较小的内存开销。
 *
 * <p>通过{@link #getRowLimit}和{@link #getColumnLimit}可以限制行列数，xlsx格式默认的行列限制分别是
 * {@code 1048576}和{@code 65536}，可以重置行限制以提前触发分页，例：{@code getRowLimit}返回{@code 10000}，
 * 则工作表Worksheet每1万行进行一次分页</p>
 *
 * @author guanquan.wang at 2019-04-22 17:23
 * @see org.ttzero.excel.entity.e7.XMLWorksheetWriter
 * @see org.ttzero.excel.entity.csv.CSVWorksheetWriter
 */
public interface IWorksheetWriter extends Closeable, Cloneable, Storable {

    /**
     * 获取最大行上限，随输出格式而定
     *
     * @return 行最大上限值
     */
    int getRowLimit();

    /**
     * 获取最大列上限，随输出格式而定
     *
     * @return 列最大上限值
     */
    int getColumnLimit();

    /**
     * 数据输出到指定位置
     *
     * @param path     保存位置
     * @param supplier 数据提供方
     * @throws IOException if I/O error occur
     * @deprecated 未使用，即将移除
     */
    @Deprecated
    default void writeTo(Path path, Supplier<RowBlock> supplier) throws IOException {
        throw new UnsupportedOperationException();
    }

    /**
     * 设置工作表
     *
     * @param sheet 工作表{@link Sheet}
     * @return 当前输出协议
     */
    IWorksheetWriter setWorksheet(Sheet sheet);

    /**
     * 复制工作表输出协议
     *
     * @return IWorksheetWriter
     */
    IWorksheetWriter clone();

    /**
     * 获取扩展名，随输出协议而定
     *
     * @return 扩展名
     */
    String getFileSuffix();

    /**
     * 添加图片
     *
     * @param picture 可写图片
     * @throws IOException if I/O error occur
     */
    default void writePicture(Picture picture) throws IOException { }

    /**
     * 写行数据
     *
     * @param rowBlock 行块
     * @throws IOException if I/O error occur
     */
    void writeData(RowBlock rowBlock) throws IOException;

    /**
     * 获取数据样式转换器，可以根据不同输出协议制定转换器
     *
     * @return 数据样式转换器
     */
    default ICellValueAndStyle getCellValueAndStyle() {
        return new XMLCellValueAndStyle();
    }

    /**
     * 判断是否为{@link java.util.Date}类型
     *
     * @param clazz the type
     * @return bool
     */
    static boolean isDate(Class<?> clazz) {
        return clazz == java.util.Date.class
            || clazz == java.sql.Date.class;
    }

    /**
     * 判断是否为{@link java.sql.Timestamp}类型
     *
     * @param clazz the type
     * @return bool
     */
    static boolean isDateTime(Class<?> clazz) {
        return clazz == java.sql.Timestamp.class;
    }

    /**
     * 判断是否为{@code int, char, byte or short}或包装类型
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
     * 判断是否为{@code short}或{@link Short}类型
     *
     * @param clazz the class
     * @return boolean value
     */
    static boolean isShort(Class<?> clazz) {
        return clazz == short.class || clazz == Short.class;
    }

    /**
     * 判断是否为{@code long}或{@link Long}类型
     *
     * @param clazz the type
     * @return bool
     */
    static boolean isLong(Class<?> clazz) {
        return clazz == long.class || clazz == Long.class;
    }

    /**
     * 判断是否为单精度浮点类型
     *
     * @param clazz the type
     * @return bool
     */
    static boolean isFloat(Class<?> clazz) {
        return clazz == float.class || clazz == Float.class;
    }

    /**
     * 判断是否为双精度浮点类型
     *
     * @param clazz the type
     * @return boolean value
     */
    static boolean isDouble(Class<?> clazz) {
        return clazz == double.class || clazz == Double.class;
    }

    /**
     * 判断是否为{@code boolean}或{@link Boolean}类型
     *
     * @param clazz the type
     * @return bool
     */
    static boolean isBool(Class<?> clazz) {
        return clazz == boolean.class || clazz == Boolean.class;
    }

    /**
     * 判断是否为{@link String}类型
     *
     * @param clazz the type
     * @return bool
     */
    static boolean isString(Class<?> clazz) {
        return clazz == String.class || clazz == CharSequence.class;
    }

    /**
     * 判断是否为{@code char} 或 {@link Character}类型
     *
     * @param clazz the type
     * @return bool
     */
    static boolean isChar(Class<?> clazz) {
        return clazz == char.class || clazz == Character.class;
    }

    /**
     * 判断是否为{@link BigDecimal}类型
     *
     * @param clazz the type
     * @return bool
     */
    static boolean isBigDecimal(Class<?> clazz) {
        return clazz == BigDecimal.class;
    }

    /**
     * 判断是否为{@link LocalDate}类型
     *
     * @param clazz the type
     * @return bool
     */
    static boolean isLocalDate(Class<?> clazz) {
        return clazz == LocalDate.class;
    }

    /**
     * 判断是否为{@link LocalDateTime}类型
     *
     * @param clazz the type
     * @return bool
     */
    static boolean isLocalDateTime(Class<?> clazz) {
        return clazz == LocalDateTime.class;
    }

    /**
     * 判断是否为{@link java.sql.Time}类型
     *
     * @param clazz the type
     * @return bool
     */
    static boolean isTime(Class<?> clazz) {
        return clazz == java.sql.Time.class;
    }

    /**
     * 判断是否为{@link java.time.LocalTime}类型
     *
     * @param clazz the type
     * @return 御前
     */
    static boolean isLocalTime(Class<?> clazz) {
        return clazz == java.time.LocalTime.class;
    }

}

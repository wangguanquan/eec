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

package org.ttzero.excel.annotation;

import org.ttzero.excel.entity.Sheet;

import java.lang.annotation.Documented;
import java.lang.annotation.ElementType;
import java.lang.annotation.Inherited;
import java.lang.annotation.Repeatable;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Specify export information
 *
 * @author guanquan.wang at 2019-06-21 09:53
 */
@Target({ ElementType.FIELD, ElementType.METHOD })
@Retention(RetentionPolicy.RUNTIME)
@Inherited
@Documented
@Repeatable(ExcelColumns.class)
public @interface ExcelColumn {
    /**
     * Title of column
     *
     * @return the title of column
     */
    String value() default "";

    /**
     * Share body string
     *
     * @return true if shared
     */
    boolean share() default false;

    /**
     * Specify a comment in header column
     * <p>
     * If this annotation appears with {@code @HeaderComment} at
     * the same time, the independent {@code @HeaderComment} annotation
     * takes precedence
     * <p>
     * NOTE: This attribute only affects the header line
     *
     * @return a header {@link HeaderComment}
     */
    HeaderComment comment() default @HeaderComment;

    /**
     * Specify the cell format.
     * <p>
     * It only supports the format specified by Office excel, please refer
     * to {@link org.ttzero.excel.entity.style.NumFmt}.
     * <p>
     * If you are not sure whether the format is correct, please open
     * Office excel{@code >} Format cell{@code >} Custom to debug the
     * custom format here.
     * <p>
     * Note: It only used on Number or Date(include Timestamp, Time and java.time.*) field.
     *
     * @return the data format string
     */
    String format() default "";

    /**
     * Wrap text in a cell
     * <p>
     * Microsoft Excel can wrap text so it appears on multiple lines in a cell.
     * You can format the cell so the text wraps automatically, or enter a manual line break.
     *
     * @return true if
     */
    boolean wrapText() default false;

    /**
     * Specify the column index(zero base), Range from {@code 0} to {@code 16383} include {@code 16383}
     * <p>
     * The column set by colIndex is an absolute position. For example,
     * if {@code colIndex=100}, this column must be placed at the {@code "CV"} position
     *
     * @return -1 means unset
     */
    int colIndex() default -1;

    /**
     * If {@link Sheet#autoWidth()} is {@code true}, The column width take the minimum of `width` and `maxWidth`,
     * otherwise the column width use `maxWidth` directly as the column width
     *
     * @return max cell width, negative number means unset
     */
    double maxWidth() default -1D;

    /**
     * Hidden current column
     * <p>
     * Only set the column to hide, the data will still be written,
     * you can right-click to "un-hide" to display in file
     *
     * @return true: hidden otherwise show
     */
    boolean hide() default false;
}

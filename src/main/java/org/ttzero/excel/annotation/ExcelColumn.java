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
 * 指定Excel列属性，用于设置Title、列宽等常用属性
 *
 * <p>处于信息安全考虑，EEC只会导出标有{@code ExcelColumn}注解的属性和方法，使用{@link org.ttzero.excel.entity.Workbook#forceExport()}
 * 可以绕过此限制强制导出所有字段，但并不建议这么做，后续Java Bean被其它人添加了敏感字段则会在无预警的情况下被导出
 * 导致信息泄露，本工具不会对此类安全事故负责。</p>
 *
 * <p>多个{@code ExcelColumn}注解组合可以实现多行表头，{@link #value()}相同的行或列会自动合并。
 * 默认情况下导出顺序是按照Bean定义顺序，也可以通过{@link #colIndex()}指定列顺序，对于多行表头
 * 必须将该属性定义在最底层的{@code ExcelColumn}注解上。</p>
 *
 * <p>建议：不建议使用{@code ExcelColumn}注解来导出Excel，开发者推荐手动指定{@link org.ttzero.excel.entity.Column}
 * 后者不会破坏Java对象，且有更丰富的属性和转换器</p>
 *
 * <p>参考文档:</p>
 * <p><a href="https://github.com/wangguanquan/eec/wiki/7-%E8%AE%BE%E7%BD%AE%E5%A4%9A%E8%A1%8C%E8%A1%A8%E5%A4%B4">设置多行表头</a></p>
 * <p><a href="https://github.com/wangguanquan/eec/wiki/5-%E6%8C%87%E5%AE%9A%E5%AF%BC%E5%87%BA%E6%97%B6%E7%9A%84%E5%88%97%E9%A1%BA%E5%BA%8F%E5%92%8C%E4%BD%8D%E7%BD%AE">指定导出时的列顺序和位置</a></p>
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
     * 设置列名，如果未指定则默认使用字段或者方法名做为列名
     *
     * @return 列名
     */
    String value() default "";

    /**
     * 设置字符串共享
     *
     * <p>EEC默认使用{@code inline}模式输出字符串，也就是将字符串直接写到每个Cell里并不共享。对于某些枚举值的列使用
     * 字符串共享将会起到压缩目的，比如"姓别"列只会有“男”，“女”和“未知”三种值。</p>
     *
     * <p>共享字符串会将值写入一个公共区域，xlsx格式保存在{@code sharedStrings.xml}文件中，整个Workbook的
     * 所有Worksheet共享</p>
     *
     * @see org.ttzero.excel.entity.SharedStrings
     * @return true: 共享，false: 直写（默认）
     */
    boolean share() default false;

    /**
     * 设置表头“批注”
     *
     * <p>如果此注释与{@link HeaderComment}同时出现时，则独立的HeaderComment注释优先</p>
     *
     * <p>注意: 该注解只作用于表头</p>
     *
     * @return 表头“批注”
     */
    HeaderComment comment() default @HeaderComment;

    /**
     * 设置单元格格式
     *
     * <p>只支持Office excel指定的格式，请参阅{@link org.ttzero.excel.entity.style.NumFmt}.
     * 如果不知道格式是否有效可以先在Office里调试，然后将调试好的字符串复制过来即可，</p>
     *
     * <p>注意: 此属性只作用于数字或者日期(包含 Timestamp, Time and java.time.*).</p>
     *
     * <pre>
     * &#x40;ExcelColumn(format = "yyyy-mm-dd hh:mm:ss") // 日期格式化 2019-06-21 09:53:21
     * &#x40;ExcelColumn(format = "#,##0.00") // 数字格式化 13,541.00
     * </pre>
     *
     * @return 格式化
     */
    String format() default "";

    /**
     * 单元格自动换行
     *
     * <p>Microsoft Excel可以将文本换行，使其显示在单元格中的多行中。以下两种情况将自动换行，一是字符串长度超过列宽，二是字符串包含"回车"符</p>
     *
     * @return true: 自动换行 false: 不换行（默认）
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
     * If {@link Sheet#autoSize()} is {@code true} (exclusive {@code MEDIA} type), The column width take the minimum of `width` and `maxWidth`,
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

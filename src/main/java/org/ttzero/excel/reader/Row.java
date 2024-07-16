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

package org.ttzero.excel.reader;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.ttzero.excel.entity.style.Styles;
import org.ttzero.excel.util.StringUtil;

import java.lang.reflect.InvocationTargetException;
import java.math.BigDecimal;
import java.sql.Timestamp;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.util.Arrays;
import java.util.Collections;
import java.util.Date;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.StringJoiner;

import static org.ttzero.excel.reader.Cell.BLANK;
import static org.ttzero.excel.reader.Cell.BOOL;
import static org.ttzero.excel.reader.Cell.CHARACTER;
import static org.ttzero.excel.reader.Cell.DATE;
import static org.ttzero.excel.reader.Cell.DATETIME;
import static org.ttzero.excel.reader.Cell.DECIMAL;
import static org.ttzero.excel.reader.Cell.DOUBLE;
import static org.ttzero.excel.reader.Cell.EMPTY_TAG;
import static org.ttzero.excel.reader.Cell.INLINESTR;
import static org.ttzero.excel.reader.Cell.LONG;
import static org.ttzero.excel.reader.Cell.NUMERIC;
import static org.ttzero.excel.reader.Cell.SST;
import static org.ttzero.excel.reader.Cell.TIME;
import static org.ttzero.excel.reader.Cell.UNALLOCATED;
import static org.ttzero.excel.reader.Cell.UNALLOCATED_CELL;
import static org.ttzero.excel.util.DateUtil.toDate;
import static org.ttzero.excel.util.DateUtil.toLocalDate;
import static org.ttzero.excel.util.DateUtil.toLocalDateTime;
import static org.ttzero.excel.util.DateUtil.toLocalTime;
import static org.ttzero.excel.util.DateUtil.toTime;
import static org.ttzero.excel.util.DateUtil.toTimestamp;
import static org.ttzero.excel.util.StringUtil.EMPTY;
import static org.ttzero.excel.util.StringUtil.isNotBlank;
import static org.ttzero.excel.util.StringUtil.isNotEmpty;

/**
 * 行数据，每行数据都包含0个到多个单元格{@link Cell}，无论是获取单元格的数据还是样式都是通过本类实现，
 * {@link Cell}对象并不提供任何获取信息的方法，{@code Row}除了提供最原始的{@link #getInt}，{@link #getString}
 * 等方法外还能调用{@link #to}和{@link #too}方法将行转为指定对象，{@code to}方法和{@code too}的区别在于前者每一行都会
 * 创建一个独立的对象而后者是内存共享的，如果需要使用数组或集合类收集对象则需要使用{@code to}方法，流式one-by-one的场景建议
 * 使用{@code too}方法。
 *
 * <p>使用{@code to}和{@code too}方法转对象都有一个前提，那就是所转对象的属性或set方法必须使用{@link org.ttzero.excel.annotation.ExcelColumn}注解，
 * 通过表头行上的文本与&#x40;ExcelColumn注解的{@code value}值进行匹配，如果使用{@link Sheet#forceImport}强制匹配时
 * 无&#x40;ExcelColumn注解的字段将会按照字段名进行匹配，除了按照表头文本匹配外还支持列索引匹配</p>
 *
 * <p>少量数据也可以使用{@link #toMap}方法将行数据转为字典类型，为了保证列顺序返回的Map方法为{@link LinkedHashMap}，
 * 字典的Key为表头文本，Value为单元格的值，多行表头按照{@code 行1:行2}进行拼接，参考{@link Sheet#getHeader}文档。</p>
 *
 * <p>{@code Row}还提供了{@link #isEmpty}和{@link #isBlank}两个空判断以及{@link #getFirstColumnIndex}和{@link #getLastColumnIndex}
 * 两个列索引方法。{@code isEmpty}仅判断单元格起始下标是否小于结束下标，至于单元格里是否有值或样式则不在判断之内，{@code isBlank}则会
 * 判断每个单元格的值是否为{@code blank}，{@code blank}的标准为字符串{@code null}或{@code “”}其余类型{@code null}</p>
 *
 * @author guanquan.wang at 2019-04-17 11:08
 */
public class Row {
    protected final Logger LOGGER = LoggerFactory.getLogger(getClass());
    // Index to row
    protected int index = -1;
    // Index to first column (zero base, inclusive)
    protected int fc = 0;
    // Index to last column (zero base, exclusive)
    protected int lc = -1;
    // Share cell
    protected Cell[] cells;
    /**
     * The Shared String Table
     */
    protected SharedStrings sst;
    // The header row
    protected HeaderRow hr;
    protected boolean unknownLength;

    // Cache formulas
    protected PreCalc[] sharedCalc;

    /**
     * The global styles
     */
    protected Styles styles;

    /**
     * 获取行号，与你打开Excel文件看到的一样从1开始
     *
     * @return 行号
     */
    public int getRowNum() {
        return index;
    }

    /**
     * 获取首列下标 (zero base)
     *
     * @return 首列下标
     */
    public int getFirstColumnIndex() {
        return fc;
    }

    /**
     * 获取尾列下标 (zero base)
     *
     * @return 尾列下标
     */
    public int getLastColumnIndex() {
        return lc;
    }

    /**
     * Returns a global {@link Styles}
     *
     * @return a style entry
     */
    public Styles getStyles() {
        return styles;
    }

    /**
     * 判断单行无有效单元格，仅空Tag&lt;row /&gt;时返回{@code true}
     *
     * @return 未实例化行时返回{@code true}
     */
    public boolean isEmpty() {
        return lc - fc <= 0;
    }

    /**
     * 判断单行是否包含有效单元格，包含任意实例化单元格时返回{@code true}
     *
     * @return 包含任意实例化单元格时返回{@code true}
     * @see java.util.function.Predicate
     * @see #isEmpty()
     */
    public boolean nonEmpty() {
        return lc > fc;
    }

    /**
     * 判断单行的所有单元格是否为空，所有单元格无值或空字符串时返回{@code true}
     *
     * @return 所有单元格无值或空字符串时返回{@code true}
     */
    public boolean isBlank() {
        if (lc > fc) {
            for (int i = fc; i < lc; i++) {
                Cell c = cells[i];
                if (!isBlank(c)) return false;
            }
        }
        return true;
    }

    /**
     * 判断单行是否包含值，任意单元格有值且不为空字符串时返回{@code true}
     *
     * @return 任意单元格有值且不为空字符串时返回{@code true}
     * @see java.util.function.Predicate
     * @see #isBlank()
     */
    public boolean nonBlank() {
        return !isBlank();
    }

    /**
     * 获取行高，仅{@code XMLFullRow}支持
     *
     * @return 当{@code customHeight=1}时返回自定义行高否则返回{@code null}
     */
    public Double getHeight() {
        return null;
    }

    /**
     * 获取当前行是否隐藏，仅{@code XMLFullRow}支持
     *
     * @return {@code true}行隐藏
     */
    public boolean isHidden() {
        return false;
    }

    /**
     * Check the cell ranges,
     *
     * @param index the index
     * @exception IndexOutOfBoundsException If the specified {@code index}
     * argument is negative
     */
    protected void rangeCheck(int index) {
        if (index < 0)
            throw new IndexOutOfBoundsException("Index: " + index + " is negative.");
    }

    /**
     * 获取单元格{@link Cell}，获取到单元格后可用于后续取值或样式
     *
     * @param i 单元格列索引
     * @return 单元格 {@link Cell}
     * @throws IndexOutOfBoundsException 单元络索引为负数时抛此异常
     */
    public Cell getCell(int i) {
        rangeCheck(i);
        return i < lc ? cells[i] : UNALLOCATED_CELL;
    }

    /**
     * 获取单元格{@link Cell}，获取到单元格后可用于后续取值或样式，如果查找不到时则返回一个空的单元格
     *
     * @param name 表头
     * @return 单元格 {@link Cell}
     */
    public Cell getCell(String name) {
        int i;
        return hr != null && (i = hr.getIndex(name)) >= 0 && i < lc ? cells[i] : UNALLOCATED_CELL;
    }

    /**
     * 将当前行转为表头行
     *
     * @return 表头行
     */
    public HeaderRow asHeader() {
        return new HeaderRow().with(this);
    }

    /**
     * 设置表头
     *
     * @param hr {@link HeaderRow}表头
     * @return 当前行
     */
    public Row setHeader(HeaderRow hr) {
        this.hr = hr;
        return this;
    }

    /**
     * 获取表头
     *
     * @return {@link HeaderRow}表头
     */
    public HeaderRow getHeader() {
        return hr;
    }

    /**
     * 获取Shared String Table
     *
     * @return Shared String Table
     */
    public SharedStrings getSharedStrings() {
        return sst;
    }

    /**
     * 获取单元格的值并转为{@code Boolean}类型，对于非布尔类型则兼容C语言的{@code bool}判断
     *
     * @param columnIndex 单元格索引
     * @return {@code numeric}类型非{@code 0}为{@code true}其余为{@code false}，
     * {@code string}类型文本值为{@code "true"}则为{@code true}，单元格为空或未实例化时返回{@code null}
     */
    public Boolean getBoolean(int columnIndex) {
        Cell c = getCell(columnIndex);
        return getBoolean(c);
    }

    /**
     * 获取单元格的值并转为{@code Boolean}类型，对于非布尔类型则兼容C语言的{@code bool}判断
     *
     * @param columnName 列名
     * @return {@code numeric}类型非{@code 0}为{@code true}其余为{@code false}，
     * {@code string}类型文本值为{@code "true"}则为{@code true}，单元格为空或未实例化时返回{@code null}
     */
    public Boolean getBoolean(String columnName) {
        Cell c = getCell(columnName);
        return getBoolean(c);
    }

    /**
     * 获取单元格的值并转为{@code Boolean}类型，对于非布尔类型则兼容C语言的{@code bool}判断
     *
     * @param c 单元格{@link Cell}
     * @return {@code numeric}类型非{@code 0}为{@code true}其余为{@code false}，
     * {@code string}类型文本值为{@code "true"}则为{@code true}，单元格为空或未实例化时返回{@code null}
     */
    public Boolean getBoolean(Cell c) {
        boolean v;
        switch (c.t) {
            case BOOL       : v = c.boolVal;                                 break;
            case NUMERIC    : v = c.intVal != 0;                             break;
            case LONG       : v = c.longVal != 0L;                           break;
            case SST        : if (c.stringVal == null) c.setString(sst.get(c.intVal)); // @Mark:=>There is no missing `break`, this is normal logic here
            case INLINESTR  : v = "true".equalsIgnoreCase(c.stringVal);      break;
            case DECIMAL    : v = c.decimal.compareTo(BigDecimal.ZERO) != 0; break;
            case DOUBLE     : v = c.doubleVal != .0D;                        break;
            case BLANK      :
            case EMPTY_TAG  :
            case UNALLOCATED: return null;
            default         : v = false;
        }
        return v;
    }

    /**
     * 获取单元格的值并转为{@code Byte}类型
     *
     * @param columnIndex 列索引
     * @return {@code numeric}类型强制转为{@code byte}，其余类型返回{@code null}
     */
    public Byte getByte(int columnIndex) {
        Cell c = getCell(columnIndex);
        return getByte(c);
    }

    /**
     * 获取单元格的值并转为{@code Byte}类型
     *
     * @param columnName 列名
     * @return {@code numeric}类型强制转为{@code byte}，其余类型返回{@code null}
     */
    public Byte getByte(String columnName) {
        Cell c = getCell(columnName);
        return getByte(c);
    }

    /**
     * 获取单元格的值并转为{@code Byte}类型
     *
     * @param c 单元格{@link Cell}
     * @return {@code numeric}类型强制转为{@code byte}，其余类型返回{@code null}
     */
    public Byte getByte(Cell c) {
        byte b = 0;
        switch (c.t) {
            case NUMERIC    : b |= c.intVal;                            break;
            case LONG       : b |= c.longVal;                           break;
            case DECIMAL    : b = c.decimal.byteValue();                break;
            case DOUBLE     : b |= (int) c.doubleVal;                   break;
            case BOOL       : b |= c.boolVal ? 1 : 0;                   break;
            default         : return null;
        }
        return b;
    }

    /**
     * 获取单元格的值并转为{@code Character}类型
     *
     * @param columnIndex 列索引
     * @return {@code numeric}类型强制转为{@code char}，{@code string}类型取第一个字符，其余类型返回{@code null}
     */
    public Character getChar(int columnIndex) {
        Cell c = getCell(columnIndex);
        return getChar(c);
    }

    /**
     * 获取单元格的值并转为{@code Character}类型
     *
     * @param columnName 列名
     * @return {@code numeric}类型强制转为{@code char}，{@code string}类型取第一个字符，其余类型返回{@code null}
     */
    public Character getChar(String columnName) {
        Cell c = getCell(columnName);
        return getChar(c);
    }

    /**
     * 获取单元格的值并转为{@code Character}类型
     *
     * @param c 单元格{@link Cell}
     * @return {@code numeric}类型强制转为{@code char}，{@code string}类型取第一个字符，其余类型返回{@code null}
     */
    public Character getChar(Cell c) {
        char cc = 0;
        switch (c.t) {
            case SST        : if (c.stringVal == null) c.setString(sst.get(c.intVal)); // @Mark:=>There is no missing `break`, this is normal logic here
            case INLINESTR  : if (isNotEmpty(c.stringVal)) cc = c.stringVal.charAt(0); break;
            case NUMERIC    : cc |= c.intVal;                           break;
            case LONG       : cc |= c.longVal;                          break;
            case BOOL       : cc |= c.boolVal ? 1 : 0;                  break;
            case DECIMAL    : cc |= c.decimal.intValue();               break;
            case DOUBLE     : cc |= (int) c.doubleVal;                  break;
            default         : return null;
        }
        return cc;
    }

    /**
     * 获取单元格的值并转为{@code Short}类型
     *
     * @param columnIndex 列索引
     * @return {@code numeric}和{@code string}类型能强转为{@code short}，其余类型返回{@code null}
     */
    public Short getShort(int columnIndex) {
        Cell c = getCell(columnIndex);
        return getShort(c);
    }

    /**
     * 获取单元格的值并转为{@code Short}类型
     *
     * @param columnName 列名
     * @return {@code numeric}和{@code string}类型能强转为{@code short}，其余类型返回{@code null}
     */
    public Short getShort(String columnName) {
        Cell c = getCell(columnName);
        return getShort(c);
    }

    /**
     * 获取单元格的值并转为{@code Short}类型
     *
     * @param c 单元格{@link Cell}
     * @return {@code numeric}和{@code string}类型能强转为{@code short}，其余类型返回{@code null}
     */
    public Short getShort(Cell c) {
        short s = 0;
        switch (c.t) {
            case NUMERIC    : s |= c.intVal;                            break;
            case LONG       : s |= c.longVal;                           break;
            case DECIMAL    : s = c.decimal.shortValue();               break;
            case DOUBLE     : s |= (int) c.doubleVal;                   break;
            case SST        : if (c.stringVal == null) c.setString(sst.get(c.intVal)); // @Mark:=>There is no missing `break`, this is normal logic here
            case INLINESTR  :
                if (StringUtil.isEmpty(c.stringVal)) return null;
                String ss = c.stringVal.trim();
                int t = testNumberType(ss.toCharArray(), 0, ss.length());
                switch (t) {
                    case 1  : s |= Integer.parseInt(ss);                break;
                    case 2  : s |= Long.parseLong(ss);                  break;
                    case 3  : s = (short) Double.parseDouble(ss);       break;
                    case 0  : return null;
                    default : throw new NumberFormatException("For input string: \"" + c.stringVal + "\"");
                }                                                       break;
            case BOOL       : s |= c.boolVal ? 1 : 0;                   break;
            default         : return null;
        }
        return s;
    }

    /**
     * 获取单元格的值并转为{@code Integer}类型
     *
     * @param columnIndex 单元格索引
     * @return {@code numeric}和{@code string}类型能强转为{@code Integer}，其余类型返回{@code null}
     */
    public Integer getInt(int columnIndex) {
        Cell c = getCell(columnIndex);
        return getInt(c);
    }

    /**
     * 获取单元格的值并转为{@code Integer}类型
     *
     * @param columnName 列索引
     * @return {@code numeric}和{@code string}类型能强转为{@code Integer}，其余类型返回{@code null}
     */
    public Integer getInt(String columnName) {
        Cell c = getCell(columnName);
        return getInt(c);
    }

    /**
     * 获取单元格的值并转为{@code Integer}类型
     *
     * @param c 单元格{@link Cell}
     * @return {@code numeric}和{@code string}类型能强转为{@code Integer}，其余类型返回{@code null}
     */
    public Integer getInt(Cell c) {
        int n = 0;
        switch (c.t) {
            case NUMERIC    : n = c.intVal;                             break;
            case LONG       : n = (int) c.longVal;                      break;
            case DECIMAL    : n = c.decimal.intValue();                 break;
            case DOUBLE     : n = (int) c.doubleVal;                    break;
            case SST        : if (c.stringVal == null) c.setString(sst.get(c.intVal)); // @Mark:=>There is no missing `break`, this is normal logic here
            case INLINESTR  :
                if (StringUtil.isEmpty(c.stringVal)) return null;
                String ss = c.stringVal.trim();
                int t = testNumberType(ss.toCharArray(), 0, ss.length());
                switch (t) {
                    case 1  : n = Integer.parseInt(ss);                 break;
                    case 2  : n |= Long.parseLong(ss);                  break;
                    case 3  : n = (int) Double.parseDouble(ss);         break;
                    case 0  : return null;
                    default : throw new NumberFormatException("For input string: \"" + c.stringVal + "\"");
                }                                                       break;
            case BOOL       : n = c.boolVal ? 1 : 0;                    break;
            default         : return null;
        }
        return n;
    }

    /**
     * 获取单元格的值并转为{@code Long}类型
     *
     * @param columnIndex 单元格索引
     * @return {@code numeric}和{@code string}类型能强转为{@code Long}，其余类型返回{@code null}
     */
    public Long getLong(int columnIndex) {
        Cell c = getCell(columnIndex);
        return getLong(c);
    }

    /**
     * 获取单元格的值并转为{@code Long}类型
     *
     * @param columnName 列名
     * @return {@code numeric}和{@code string}类型能强转为{@code Long}，其余类型返回{@code null}
     */
    public Long getLong(String columnName) {
        Cell c = getCell(columnName);
        return getLong(c);
    }

    /**
     * 获取单元格的值并转为{@code Long}类型
     *
     * @param c 单元格{@link Cell}
     * @return {@code numeric}和{@code string}类型能强转为{@code Long}，其余类型返回{@code null}
     */
    public Long getLong(Cell c) {
        long l;
        switch (c.t) {
            case LONG       : l = c.longVal;                            break;
            case NUMERIC    : l = c.intVal;                             break;
            case DECIMAL    : l = c.decimal.longValue();                break;
            case DOUBLE     : l = (long) c.doubleVal;                   break;
            case SST        : if (c.stringVal == null) c.setString(sst.get(c.intVal)); // @Mark:=>There is no missing `break`, this is normal logic here
            case INLINESTR  :
                if (StringUtil.isEmpty(c.stringVal)) return null;
                String ss = c.stringVal.trim();
                int t = testNumberType(ss.toCharArray(), 0, ss.length());
                switch (t) {
                    case 1  :
                    case 2  : l = Long.parseLong(ss);                   break;
                    case 3  : l = (long) Double.parseDouble(ss);        break;
                    case 0  : return null;
                    default : throw new NumberFormatException("For input string: \"" + c.stringVal + "\"");
                }                                                       break;
            case BOOL       : l = c.boolVal ? 1L : 0L;                  break;
            default         : return null;
        }
        return l;
    }

    /**
     * 获取单元格的值并转为{@code String}类型
     *
     * @param columnIndex 单元格索引
     * @return 单元格有值时强转为字符串否则返回{@code null}
     */
    public String getString(int columnIndex) {
        Cell c = getCell(columnIndex);
        return getString(c);
    }

    /**
     * 获取单元格的值并转为{@code String}类型
     *
     * @param columnName 列名
     * @return 单元格有值时强转为字符串否则返回{@code null}
     */
    public String getString(String columnName) {
        Cell c = getCell(columnName);
        return getString(c);
    }

    /**
     * 获取单元格的值并转为{@code String}类型
     *
     * @param c 单元格{@link Cell}
     * @return 单元格有值时强转为字符串否则返回{@code null}
     */
    public String getString(Cell c) {
        String s;
        switch (c.t) {
            case SST        : if (c.stringVal == null) c.setString(sst.get(c.intVal)); // @Mark:=>There is no missing `break`, this is normal logic here
            case INLINESTR  : s = c.stringVal;                          break;
            case BLANK      :
            case EMPTY_TAG  :
            case UNALLOCATED: s = null;                                 break;
            case LONG       : s = String.valueOf(c.longVal);            break;
            case NUMERIC    : s = String.valueOf(c.intVal);             break;
            case DECIMAL    : s = c.decimal.toString();                 break;
            case DOUBLE     : s = String.valueOf(c.doubleVal);          break;
            case BOOL       : s = c.boolVal ? "true" : "false";         break;
            default         : s = c.stringVal;
        }
        return s;
    }

    /**
     * 获取单元格的值并转为{@code Float}类型
     *
     * @param columnIndex 单元格索引
     * @return 单元格有值时强转为{@code Float}否则返回{@code null}，此接口可能抛{@code NumberFormatException}异常
     */
    public Float getFloat(int columnIndex) {
        Double d = getDouble(columnIndex);
        return d != null ? Float.valueOf(d.toString()) : null;
    }

    /**
     * 获取单元格的值并转为{@code Float}类型
     *
     * @param columnName 列名
     * @return 单元格有值时强转为{@code Float}否则返回{@code null}，此接口可能抛{@code NumberFormatException}异常
     */
    public Float getFloat(String columnName) {
        Double d = getDouble(columnName);
        return d != null ? Float.valueOf(d.toString()) : null;
    }

    /**
     * 获取单元格的值并转为{@code Float}类型
     *
     * @param c 单元格{@link Cell}
     * @return 单元格有值时强转为{@code Float}否则返回{@code null}，此接口可能抛{@code NumberFormatException}异常
     */
    public Float getFloat(Cell c) {
        Double d = getDouble(c);
        return d != null ? Float.valueOf(d.toString()) : null;
    }

    /**
     * 获取单元格的值并转为{@code Double}类型
     *
     * @param columnIndex 单元格索引
     * @return 单元格有值时强转为{@code Double}否则返回{@code null}，此接口可能抛{@code NumberFormatException}异常
     */
    public Double getDouble(int columnIndex) {
        Cell c = getCell(columnIndex);
        return getDouble(c);
    }

    /**
     * 获取单元格的值并转为{@code Double}类型
     *
     * @param columnName 列名
     * @return 单元格有值时强转为{@code Double}否则返回{@code null}，此接口可能抛{@code NumberFormatException}异常
     */
    public Double getDouble(String columnName) {
        Cell c = getCell(columnName);
        return getDouble(c);
    }

    /**
     * 获取单元格的值并转为{@code Double}类型
     *
     * @param c 单元格{@link Cell}
     * @return 单元格有值时强转为{@code Double}否则返回{@code null}，此接口可能抛{@code NumberFormatException}异常
     */
    public Double getDouble(Cell c) {
        double d;
        switch (c.t) {
            case DECIMAL    : d = c.decimal.doubleValue();              break;
            case DOUBLE     : d = c.doubleVal;                          break;
            case NUMERIC    : d = c.intVal;                             break;
            case LONG       : d = c.longVal;                            break;
            case SST        : if (c.stringVal == null) c.setString(sst.get(c.intVal)); // @Mark:=>There is no missing `break`, this is normal logic here
            case INLINESTR  :
                if (isNotBlank(c.stringVal)) d = Double.parseDouble(c.stringVal.trim());
                else return null;                                       break;
            default         : return null;
        }
        return d;
    }

    /**
     * 获取单元格的值并转为{@code java.math.BigDecimal}类型
     *
     * @param columnIndex 单元格索引
     * @return 单元格有值时强转为{@code java.math.BigDecimal}否则返回{@code null}，此接口可能抛{@code NumberFormatException}异常
     */
    public BigDecimal getDecimal(int columnIndex) {
        Cell c = getCell(columnIndex);
        return getDecimal(c);
    }

    /**
     * 获取单元格的值并转为{@code java.math.BigDecimal}类型
     *
     * @param columnName 列名
     * @return 单元格有值时强转为{@code java.math.BigDecimal}否则返回{@code null}，此接口可能抛{@code NumberFormatException}异常
     */
    public BigDecimal getDecimal(String columnName) {
        Cell c = getCell(columnName);
        return getDecimal(c);
    }

    /**
     * 获取单元格的值并转为{@code java.math.BigDecimal}类型
     *
     * @param c 单元格{@link Cell}
     * @return 单元格有值时强转为{@code java.math.BigDecimal}否则返回{@code null}，此接口可能抛{@code NumberFormatException}异常
     */
    public BigDecimal getDecimal(Cell c) {
        BigDecimal bd;
        switch (c.t) {
            case DECIMAL    : bd = c.decimal;                            break;
            case DOUBLE     : bd = BigDecimal.valueOf(c.doubleVal);      break;
            case NUMERIC    : bd = BigDecimal.valueOf(c.intVal);         break;
            case LONG       : bd = BigDecimal.valueOf(c.longVal);        break;
            case SST        : if (c.stringVal == null) c.setString(sst.get(c.intVal)); // @Mark:=>There is no missing `break`, this is normal logic here
            case INLINESTR  : bd = isNotBlank(c.stringVal) ? new BigDecimal(c.stringVal.trim()) : null; break;
            default         : bd = null;
        }
        return bd;
    }

    /**
     * 获取单元格的值并转为{@code java.util.Date}类型，整数和小数类型将以{@code 1900-1-1}为基础进行计算，字符串将进行格式化处理
     *
     * @param columnIndex 单元格索引
     * @return 单元格有值时强转为{@code java.util.Date}否则返回{@code null}，此接口可能抛{@code NumberFormatException}异常
     */
    public Date getDate(int columnIndex) {
        Cell c = getCell(columnIndex);
        return getDate(c);
    }

    /**
     * 获取单元格的值并转为{@code java.util.Date}类型，整数和小数类型将以{@code 1900-1-1}为基础进行计算，字符串将进行格式化处理
     *
     * @param columnName 列名
     * @return 单元格有值时强转为{@code java.util.Date}否则返回{@code null}，此接口可能抛{@code NumberFormatException}异常
     */
    public Date getDate(String columnName) {
        Cell c = getCell(columnName);
        return getDate(c);
    }

    /**
     * 获取单元格的值并转为{@code java.util.Date}类型，整数和小数类型将以{@code 1900-1-1}为基础进行计算，字符串将进行格式化处理
     *
     * @param c 单元格{@link Cell}
     * @return 单元格有值时强转为{@code java.util.Date}否则返回{@code null}，此接口可能抛{@code NumberFormatException}异常
     */
    public Date getDate(Cell c) {
        Date date;
        switch (c.t) {
            case NUMERIC    : date = toDate(c.intVal);                  break;
            case DECIMAL    : date = toDate(c.decimal.doubleValue());   break;
            case DOUBLE     : date = toDate(c.doubleVal);               break;
            case SST        : if (c.stringVal == null) c.setString(sst.get(c.intVal)); // @Mark:=>There is no missing `break`, this is normal logic here
            case INLINESTR  : date = isNotBlank(c.stringVal) ? toDate(c.stringVal.trim()) : null; break;
            default         : date = null;
        }
        return date;
    }

    /**
     * 获取单元格的值并转为{@code java.sql.Timestamp}类型，整数和小数类型将以{@code 1900-1-1}为基础进行计算，字符串将进行格式化处理
     *
     * @param columnIndex 单元格索引
     * @return 单元格有值时强转为{@code java.sql.Timestamp}否则返回{@code null}，此接口可能抛{@code NumberFormatException}异常
     */
    public Timestamp getTimestamp(int columnIndex) {
        Cell c = getCell(columnIndex);
        return getTimestamp(c);
    }

    /**
     * 获取单元格的值并转为{@code java.sql.Timestamp}类型，整数和小数类型将以{@code 1900-1-1}为基础进行计算，字符串将进行格式化处理
     *
     * @param columnName 列名
     * @return 单元格有值时强转为{@code java.sql.Timestamp}否则返回{@code null}，此接口可能抛{@code NumberFormatException}异常
     */
    public Timestamp getTimestamp(String columnName) {
        Cell c = getCell(columnName);
        return getTimestamp(c);
    }

    /**
     * 获取单元格的值并转为{@code java.sql.Timestamp}类型，整数和小数类型将以{@code 1900-1-1}为基础进行计算，字符串将进行格式化处理
     *
     * @param c 单元格{@link Cell}
     * @return 单元格有值时强转为{@code java.sql.Timestamp}否则返回{@code null}，此接口可能抛{@code NumberFormatException}异常
     */
    public Timestamp getTimestamp(Cell c) {
        Timestamp ts;
        switch (c.t) {
            case NUMERIC    : ts = toTimestamp(c.intVal);                break;
            case DECIMAL    : ts = toTimestamp(c.decimal.doubleValue()); break;
            case DOUBLE     : ts = toTimestamp(c.doubleVal);             break;
            case SST        : if (c.stringVal == null) c.setString(sst.get(c.intVal)); // @Mark:=>There is no missing `break`, this is normal logic here
            case INLINESTR  : ts = isNotBlank(c.stringVal) ? toTimestamp(c.stringVal.trim()) : null; break;
            default         : ts = null;
        }
        return ts;
    }

    /**
     * 获取单元格的值并转为{@code java.sql.Time}类型，整数和小数类型将以{@code 1900-1-1}为基础进行计算，字符串将进行格式化处理
     *
     * @param columnIndex 单元格索引
     * @return 单元格有值时强转为{@code java.sql.Time}否则返回{@code null}，此接口可能抛{@code NumberFormatException}异常
     */
    public java.sql.Time getTime(int columnIndex) {
        return getTime(getCell(columnIndex));
    }

    /**
     * 获取单元格的值并转为{@code java.sql.Time}类型，整数和小数类型将以{@code 1900-1-1}为基础进行计算，字符串将进行格式化处理
     *
     * @param columnName 列名
     * @return 单元格有值时强转为{@code java.sql.Time}否则返回{@code null}，此接口可能抛{@code NumberFormatException}异常
     */
    public java.sql.Time getTime(String columnName) {
        return getTime(getCell(columnName));
    }

    /**
     * 获取单元格的值并转为{@code java.sql.Time}类型，整数和小数类型将以{@code 1900-1-1}为基础进行计算，字符串将进行格式化处理
     *
     * @param c 单元格{@link Cell}
     * @return 单元格有值时强转为{@code java.sql.Time}否则返回{@code null}，此接口可能抛{@code NumberFormatException}异常
     */
    public java.sql.Time getTime(Cell c) {
        java.sql.Time t;
        switch (c.t) {
            case DECIMAL    : t = toTime(c.decimal.doubleValue());                          break;
            case DOUBLE     : t = toTime(c.doubleVal);                                      break;
            case SST        : if (c.stringVal == null) c.setString(sst.get(c.intVal)); // @Mark:=>There is no missing `break`, this is normal logic here
            case INLINESTR  : t = isNotBlank(c.stringVal) ? toTime(c.stringVal.trim()) : null; break;
            default         : t = null;
        }
        return t;
    }

    /**
     * 获取单元格的值并转为{@code LocalDateTime}类型，整数和小数类型将以{@code 1900-1-1}为基础进行计算，字符串将进行格式化处理
     *
     * @param columnIndex 单元格索引
     * @return 单元格有值时强转为{@code LocalDateTime}否则返回{@code null}，此接口可能抛{@code NumberFormatException}异常
     */
    public LocalDateTime getLocalDateTime(int columnIndex) {
        Cell c = getCell(columnIndex);
        return getLocalDateTime(c);
    }

    /**
     * 获取单元格的值并转为{@code LocalDateTime}类型，整数和小数类型将以{@code 1900-1-1}为基础进行计算，字符串将进行格式化处理
     *
     * @param columnName 列名
     * @return 单元格有值时强转为{@code LocalDateTime}否则返回{@code null}，此接口可能抛{@code NumberFormatException}异常
     */
    public LocalDateTime getLocalDateTime(String columnName) {
        Cell c = getCell(columnName);
        return getLocalDateTime(c);
    }

    /**
     * 获取单元格的值并转为{@code LocalDateTime}类型，整数和小数类型将以{@code 1900-1-1}为基础进行计算，字符串将进行格式化处理
     *
     * @param c 单元格{@link Cell}
     * @return 单元格有值时强转为{@code LocalDateTime}否则返回{@code null}，此接口可能抛{@code NumberFormatException}异常
     */
    public LocalDateTime getLocalDateTime(Cell c) {
        LocalDateTime ldt;
        switch (c.t) {
            case NUMERIC    : ldt = toLocalDateTime(c.intVal);                              break;
            case DECIMAL    : ldt = toLocalDateTime(c.decimal.doubleValue());               break;
            case DOUBLE     : ldt = toLocalDateTime(c.doubleVal);                           break;
            case SST        : if (c.stringVal == null) c.setString(sst.get(c.intVal)); // @Mark:=>There is no missing `break`, this is normal logic here
            case INLINESTR  : ldt = isNotBlank(c.stringVal) ? toTimestamp(c.stringVal.trim()).toLocalDateTime() : null; break;
            default         : ldt = null;
        }
        return ldt;
    }

    /**
     * 获取单元格的值并转为{@code LocalDate}类型，整数和小数类型将以{@code 1900-1-1}为基础进行计算，字符串将进行格式化处理
     *
     * @param columnIndex 单元格索引
     * @return 单元格有值时强转为{@code LocalDate}否则返回{@code null}，此接口可能抛{@code NumberFormatException}异常
     */
    public LocalDate getLocalDate(int columnIndex) {
        Cell c = getCell(columnIndex);
        return getLocalDate(c);
    }

    /**
     * 获取单元格的值并转为{@code LocalDate}类型，整数和小数类型将以{@code 1900-1-1}为基础进行计算，字符串将进行格式化处理
     *
     * @param columnName 列名
     * @return 单元格有值时强转为{@code LocalDate}否则返回{@code null}，此接口可能抛{@code NumberFormatException}异常
     */
    public LocalDate getLocalDate(String columnName) {
        Cell c = getCell(columnName);
        return getLocalDate(c);
    }

    /**
     * 获取单元格的值并转为{@code LocalDate}类型，整数和小数类型将以{@code 1900-1-1}为基础进行计算，字符串将进行格式化处理
     *
     * @param c 单元格{@link Cell}
     * @return 单元格有值时强转为{@code LocalDate}否则返回{@code null}，此接口可能抛{@code NumberFormatException}异常
     */
    public LocalDate getLocalDate(Cell c) {
        LocalDate ld;
        switch (c.t) {
            case NUMERIC    : ld = toLocalDate(c.intVal);                   break;
            case DECIMAL    : ld = toLocalDate(c.decimal.intValue());       break;
            case DOUBLE     : ld = toLocalDate((int) c.doubleVal);          break;
            case SST        : if (c.stringVal == null) c.setString(sst.get(c.intVal)); // @Mark:=>There is no missing `break`, this is normal logic here
            case INLINESTR  : ld = isNotBlank(c.stringVal) ? toTimestamp(c.stringVal.trim()).toLocalDateTime().toLocalDate() : null; break;
            default         : ld = null;
        }
        return ld;
    }

    /**
     * 获取单元格的值并转为{@code LocalTime}类型，整数和小数类型将以{@code 1900-1-1}为基础进行计算，字符串将进行格式化处理
     *
     * @param columnIndex 单元格索引
     * @return 单元格有值时强转为{@code LocalTime}否则返回{@code null}，此接口可能抛{@code NumberFormatException}异常
     */
    public LocalTime getLocalTime(int columnIndex) {
        Cell c = getCell(columnIndex);
        return getLocalTime(c);
    }

    /**
     * 获取单元格的值并转为{@code LocalTime}类型，整数和小数类型将以{@code 1900-1-1}为基础进行计算，字符串将进行格式化处理
     *
     * @param columnName 列名
     * @return 单元格有值时强转为{@code LocalTime}否则返回{@code null}，此接口可能抛{@code NumberFormatException}异常
     */
    public LocalTime getLocalTime(String columnName) {
        Cell c = getCell(columnName);
        return getLocalTime(c);
    }

    /**
     * 获取单元格的值并转为{@code LocalTime}类型，整数和小数类型将以{@code 1900-1-1}为基础进行计算，字符串将进行格式化处理
     *
     * @param c 单元格{@link Cell}
     * @return 单元格有值时强转为{@code LocalTime}否则返回{@code null}，此接口可能抛{@code NumberFormatException}异常
     */
    public LocalTime getLocalTime(Cell c) {
        LocalTime lt;
        switch (c.t) {
            case NUMERIC     : lt = toLocalTime(c.intVal);                  break;
            case DECIMAL     : lt = toLocalTime(c.decimal.doubleValue());   break;
            case DOUBLE      : lt = toLocalTime(c.doubleVal);               break;
            case SST         : if (c.stringVal == null) c.setString(sst.get(c.intVal));// @Mark:=>There is no missing `break`, this is normal logic here
            case INLINESTR   :
                if (isNotBlank(c.stringVal)) {
                    c.stringVal = c.stringVal.trim();
                    // 00:00:00
                    if (c.stringVal.length() == 8 && c.stringVal.charAt(2) == ':' && c.stringVal.charAt(5) == ':') lt = toLocalTime(c.stringVal);
                    else lt = toTimestamp(c.stringVal).toLocalDateTime().toLocalTime();
                } else lt = null;
                break;
            default          : lt = null;
        }
        return lt;
    }

    /**
     * 获取单元格的公式
     *
     * @param columnIndex 列索引
     * @return 单元格含有公式时返回公式字符串，否则返回{@code null}
     */
    public String getFormula(int columnIndex) {
        return getCell(columnIndex).formula;
    }

    /**
     * 获取单元格的公式
     *
     * @param columnName 列名
     * @return 单元格含有公式时返回公式字符串，否则返回{@code null}
     */
    public String getFormula(String columnName) {
        return getCell(columnName).formula;
    }

    /**
     * 获取单元格的公式
     *
     * @param cell 单元格
     * @return 单元格含有公式时返回公式字符串，否则返回{@code null}
     */
    public String getFormula(Cell cell) {
        return cell.formula;
    }

    /**
     * 检查单元格是否包含公式
     *
     * @param columnIndex 单元格索引
     * @return 单元格含有公式时返回{@code true}否则返回{@code false}
     */
    public boolean hasFormula(int columnIndex) {
        return getCell(columnIndex).f;
    }

    /**
     * 检查单元格是否包含公式
     *
     * @param columnName 列名
     * @return 单元格含有公式时返回{@code true}否则返回{@code false}
     */
    public boolean hasFormula(String columnName) {
        return getCell(columnName).f;
    }

    /**
     * 检查单元格是否包含公式
     *
     * @param cell 单元格
     * @return 单元格含有公式时返回{@code true}否则返回{@code false}
     */
    public boolean hasFormula(Cell cell) {
        return cell.f;
    }

    /**
     * 获取单元格的数据类型
     *
     * <p>注意：这里仅是一个近似的类型，因为从原始文件中只能获取到{@code numeric}，{@code string}，
     * {@code boolean}三种类型。在解析到原始类型之后对于{@code numeric}类型会根据数字大小和格式重新设置类型，
     * 比如小于{@code 0x7fffffff}的值为{@code int}，超过这个范围时为{@code long}，如果带日期格式化则为{@code date}类型，
     * 小数也是同样处理，有日期格式化为{@code data}类型否则为{@code decimal}类型</p>
     *
     * @param columnIndex 列索引
     * @return 单元格的数据类型 {@link CellType}
     */
    public CellType getCellType(int columnIndex) {
        Cell c = getCell(columnIndex);
        return getCellType(c);
    }

    /**
     * 获取单元格的数据类型
     *
     * <p>注意：这里仅是一个近似的类型，因为从原始文件中只能获取到{@code numeric}，{@code string}，
     * {@code boolean}三种类型。在解析到原始类型之后对于{@code numeric}类型会根据数字大小和格式重新设置类型，
     * 比如小于{@code 0x7fffffff}的值为{@code int}，超过这个范围时为{@code long}，如果带日期格式化则为{@code date}类型，
     * 小数也是同样处理，有日期格式化为{@code data}类型否则为{@code decimal}类型</p>
     *
     * @param columnName 列名
     * @return 单元格的数据类型 {@link CellType}
     */
    public CellType getCellType(String columnName) {
        Cell c = getCell(columnName);
        return getCellType(c);
    }

    /**
     * 获取单元格的数据类型
     *
     * <p>注意：这里仅是一个近似的类型，因为从原始文件中只能获取到{@code numeric}，{@code string}，
     * {@code boolean}三种类型。在解析到原始类型之后对于{@code numeric}类型会根据数字大小和格式重新设置类型，
     * 比如小于{@code 0x7fffffff}的值为{@code int}，超过这个范围时为{@code long}，如果带日期格式化则为{@code date}类型，
     * 小数也是同样处理，有日期格式化为{@code data}类型否则为{@code decimal}类型</p>
     *
     * @param c 单元格{@link Cell}
     * @return 单元格的数据类型 {@link CellType}
     */
    public CellType getCellType(Cell c) {
        CellType type;
        switch (c.t) {
            case SST        :
            case INLINESTR  : type = CellType.STRING;                                                  break;
            case NUMERIC    :
            case CHARACTER  : type = !styles.fastTestDateFmt(c.xf) ? CellType.INTEGER : CellType.DATE; break;
            case LONG       : type = CellType.LONG;                                                    break;
            case DECIMAL    : type = !styles.fastTestDateFmt(c.xf) ? CellType.DECIMAL : CellType.DATE; break;
            case DOUBLE     : type = !styles.fastTestDateFmt(c.xf) ? CellType.DOUBLE : CellType.DATE;  break;
            case DATETIME   :
            case DATE       :
            case TIME       : type = CellType.DATE;                                                    break;
            case BOOL       : type = CellType.BOOLEAN;                                                 break;
            case EMPTY_TAG  :
            case BLANK      : type = CellType.BLANK;                                                   break;
            case UNALLOCATED: type = CellType.UNALLOCATED;                                             break;
            default         : type = CellType.STRING;
        }
        return type;
    }

    /**
     * 获取单元格样式值，可以拿此返回值调用{@link Styles#getBorder(int)}等方法获取具体的样式
     *
     * @param columnIndex 列索引
     * @return 样式值
     */
    public int getCellStyle(int columnIndex) {
        Cell c = getCell(columnIndex);
        return getCellStyle(c);
    }

    /**
     * 获取单元格样式值，可以拿此返回值调用{@link Styles#getBorder(int)}等方法获取具体的样式
     *
     * @param columnName 列名
     * @return 样式值
     */
    public int getCellStyle(String columnName) {
        Cell c = getCell(columnName);
        return getCellStyle(c);
    }

    /**
     * 获取单元格样式值，可以拿此返回值调用{@link Styles#getBorder(int)}等方法获取具体的样式
     *
     * @param c 单元格{@link Cell}
     * @return 样式值
     */
    public int getCellStyle(Cell c) {
        return styles.getStyleByIndex(c.xf);
    }

    /**
     * 判断单元格是否为空值
     *
     * @param columnIndex 列索引
     * @return 单元格无值或空字符串时返回{@code true}
     */
    public boolean isBlank(int columnIndex) {
        Cell c = getCell(columnIndex);
        return isBlank(c);
    }

    /**
     * 判断单元格是否为空值
     *
     * @param columnName 列名
     * @return 单元格无值或空字符串时返回{@code true}
     */
    public boolean isBlank(String columnName) {
        Cell c = getCell(columnName);
        return isBlank(c);
    }

    /**
     * 判断单元格是否为空值
     *
     * @param c 单元格{@link Cell}
     * @return 单元格无值或空字符串时返回{@code true}
     */
    public boolean isBlank(Cell c) {
        boolean blank;
        switch (c.t) {
            case SST        : if (c.stringVal == null) c.setString(sst.get(c.intVal)); // @Mark:=>There is no missing `break`, this is normal logic here
            case INLINESTR  : blank = StringUtil.isBlank(c.stringVal); break;
            case BLANK      :
            case EMPTY_TAG  :
            case UNALLOCATED: blank = true; break;
            default         : blank = false;
        }
        return blank;
    }

    /**
     * 使用{@link Sheet#bind}方法绑定类型后，使用此方法将整行数据转为指定类型&lt;T&gt;
     *
     * @param <T> 绑定的对象类型
     * @return T
     */
    @SuppressWarnings("unchecked")
    public <T> T get() {
        if (hr != null && hr.getClazz() != null) {
            T t;
            try {
                t = (T) hr.getClazz().newInstance();
                hr.put(this, t);
            } catch (InstantiationException | IllegalAccessException | InvocationTargetException e) {
                throw new UncheckedTypeException(hr.getClazz() + " new instance error.", e);
            }
            return t;
        }
//        else return (T) this;
        throw new ExcelReadException("It can only be used after binding with method `Sheet#bind`");
    }

    /**
     * 使用{@link Sheet#bind}方法绑定类型后，使用此方法将整行数据转为指定类型&lt;T&gt;，与{@link #get}的区别在于
     * 本方法返回的对象是内存共享的（只有一个对象副本）
     *
     * @param <T> 绑定的对象类型
     * @return T
     */
    public <T> T geet() {
        if (hr != null && hr.getClazz() != null) {
            T t = hr.getT();
            try {
                hr.put(this, t);
            } catch (IllegalAccessException | InvocationTargetException e) {
                throw new UncheckedTypeException("call set method error.", e);
            }
            return t;
        }
//        else return (T) this;
        throw new ExcelReadException("It can only be used after binding with method `Sheet#bind`");
    }
    /////////////////////////////To object//////////////////////////////////

    /**
     * 将行数据转为指定对象&lt;T&gt;，待转换的对象必须包含无参的构建函数且待接收的字段使用{@link org.ttzero.excel.annotation.ExcelColumn}注解，
     * 如果未指定表头时则以当前行为表头此时{@code to}方法会返回一个{@code null}对象，外部需要过滤非{@code null}对象否则会抛NPE异常。
     *
     * <p>指定对象&lt;T&gt;解析的结果会缓存到{@code HeaderRow}对象中，除非指定不同类型否则后续都将从{@code HeaderRow}中获取
     * 必要信息，这样可以提高转换性能</p>
     *
     * @param clazz 指定转换类型
     * @param <T>   强转返回对象类型
     * @return T
     */
    public <T> T to(Class<T> clazz) {
        if (hr == null) {
            hr = asHeader();
            return null;
        }
        // reset class info
        if (!hr.is(clazz)) {
            hr.setClass(clazz);
        }
        T t;
        try {
            t = clazz.newInstance();
            hr.put(this, t);
        } catch (InstantiationException | IllegalAccessException | InvocationTargetException e) {
            throw new UncheckedTypeException(clazz + " new instance error.", e);
        }
        return t;
    }

    /**
     * 与{@link #to}方法功能相同，唯一区别是{@code #too}方法返回的对象是内存共享的，所以不能将返回值收集到集合类或者数组
     *
     * @param clazz 指定转换类型
     * @param <T>   强转返回对象类型
     * @return T
     */
    public <T> T too(Class<T> clazz) {
        if (hr == null) {
            hr = asHeader();
            return null;
        }
        // reset class info
        if (!hr.is(clazz)) {
            try {
                hr.setClassOnce(clazz);
            } catch (IllegalAccessException | InstantiationException e) {
                throw new UncheckedTypeException(clazz + " new instance error.", e);
            }
        }
        T t = hr.getT();
        try {
            hr.put(this, t);
        } catch (IllegalAccessException | InvocationTargetException e) {
            throw new UncheckedTypeException("call set method error.", e);
        }
        return t;
    }

    @Override
    public String toString() {
        if (isEmpty()) return "";
        StringJoiner joiner = new StringJoiner(" | ");
        // show row number
//        joiner.add(String.valueOf(getRowNumber()));
        for (int i = fc; i < lc; i++) {
            Cell c = cells[i];
            switch (c.t) {
                case SST      : if (c.stringVal == null) c.setString(sst.get(c.intVal)); // @Mark:=>There is no missing `break`, this is normal logic here
                case INLINESTR: joiner.add(c.stringVal); break;
                case NUMERIC  :
                    if (!styles.fastTestDateFmt(c.xf)) joiner.add(String.valueOf(c.intVal));
                    else joiner.add(toLocalDate(c.intVal).toString());
                    break;
                case LONG     : joiner.add(String.valueOf(c.longVal)); break;
                case DECIMAL:
                    if (!styles.fastTestDateFmt(c.xf)) joiner.add(c.decimal.toString());
                    else if (c.decimal.compareTo(BigDecimal.ONE) > 0) joiner.add(toTimestamp(c.decimal.doubleValue()).toString());
                    else joiner.add(toLocalTime(c.decimal.doubleValue()).toString());
                    break;
                case DOUBLE:
                    if (!styles.fastTestDateFmt(c.xf)) joiner.add(String.valueOf(c.doubleVal));
                    else if (c.doubleVal > 1.0000) joiner.add(toTimestamp(c.doubleVal).toString());
                    else joiner.add(toLocalTime(c.doubleVal).toString());
                    break;
                case BLANK    :
                case EMPTY_TAG: joiner.add(EMPTY); break;
                case BOOL     : joiner.add(String.valueOf(c.boolVal)); break;
                default       : joiner.add(null);
            }
        }
        return joiner.toString();
    }

    /**
     * 将行数据转为字典类型，为保证列顺序实际类型为{@code LinkedHashMap}，如果使用{@link Sheet#dataRows}和{@link Sheet#header}
     * 指定表头则字典的Key为表头文本，Value为表头对应的列值，如果未指定表头那将以列索引做为Key，与导出指定的colIndex一样索引从{@code 0}开始。
     * 对于多行表头字典Key将以{@code 行1:行2:行n}的格式进行拼接，横向合并的单元格将自动将值复制到每一列，而纵向合并的单元格则不会复制，
     * 可以参考{@link Sheet#getHeader}方法。
     *
     * <p>关于单元格类型的特殊说明：行数据转对象时会根据对象定义进行一次类型转换，将单元格的值转为对象定义中的类型，但是转为字典时却不会有这一步
     * 逻辑，类型是根据excel中的值进行粗粒度转换，例如数字类型如果带有日期格式化则会返回一个{@code Timestamp}类型，
     * 所以最终的数据类型可能与预期有所不同</p>
     *
     * @return 按列顺序为基础的LinkedHashMap
     */
    public Map<String, Object> toMap() {
        if (isEmpty()) return Collections.emptyMap();
        boolean hasHeader = hr != null;
        // Maintain the column orders
        Map<String, Object> data = new LinkedHashMap<>(hasHeader ? Math.max(16, hr.lc - hr.fc) : 16);
        String[] names = hasHeader ? hr.names : null;
        String key;
        int from = hasHeader ? hr.fc : fc, to = hasHeader ? hr.lc : lc;
        for (int i = from; i < to; i++) {
            // Cell c = cells[i];
            Cell c = getCell(i);
            key = hasHeader ? names[i] : Integer.toString(i);
            // Ignore null key
            if (key == null) continue;
            switch (c.t) {
                case SST:
                    if (c.stringVal == null) c.setString(sst.get(c.intVal));
                    // @Mark:=>There is no missing `break`, this is normal logic here
                case INLINESTR:
                    data.put(key, c.stringVal);
                    break;
                case NUMERIC:
                    if (!styles.fastTestDateFmt(c.xf)) data.put(key, c.intVal);
                    else data.put(key, toTimestamp(c.intVal));
                    break;
                case LONG:
                    data.put(key, c.longVal);
                    break;
                case DECIMAL:
                    if (!styles.fastTestDateFmt(c.xf)) data.put(key, c.decimal);
                    else if (c.decimal.compareTo(BigDecimal.ONE) > 0) data.put(key, toTimestamp(c.decimal.doubleValue()));
                    else data.put(key, toTime(c.decimal.doubleValue()));
                    break;
                case DOUBLE:
                    if (!styles.fastTestDateFmt(c.xf)) data.put(key, c.doubleVal);
                    else if (c.doubleVal > 1.00000) data.put(key, toTimestamp(c.doubleVal));
                    else data.put(key, toTime(c.doubleVal));
                    break;
                case BLANK:
                case EMPTY_TAG:
                    data.put(key, EMPTY);
                    break;
                case BOOL:
                    data.put(key, c.boolVal);
                    break;
                default:
                    data.put(key, null);
            }
        }
        return data;
    }

    /**
     * Add function shared ref
     * <blockquote><pre>
     * 63        | Not used
     * ----------+------------
     * 42-62     | First row number
     * 28-41     | First column number
     * 8-27/14-27| Size, if axis is zero the size used 20 bits, otherwise used 14 bits
     * 2-7/2-13  | Not used
     * 0-1       | Axis, 00: range 01: y-axis 10: x-axis
     * </pre></blockquote>
     *
     * @param i the ref id
     * @param ref ref value, a range dimension string
     */
    void addRef(int i, String ref) {
        if (StringUtil.isEmpty(ref) || ref.indexOf(':') < 0)
            return;

        if (sharedCalc == null) {
            sharedCalc = new PreCalc[Math.max(10, i + 1)];
        } else if (i >= sharedCalc.length) {
            sharedCalc = Arrays.copyOf(sharedCalc, i + 10);
        }
        Dimension dim = Dimension.of(ref);

        long l = 0;
        l |= (long) (dim.firstRow & (1 << 20) - 1) << 42;
        l |= (long) (dim.firstColumn & (1 << 14) - 1) << 28;

        if (dim.firstColumn == dim.lastColumn) {
            l |= ((dim.lastRow - dim.firstRow) & (1 << 20) - 1) << 8;
            l |= (1 << 1);
        }
        else if (dim.firstRow == dim.lastRow) {
            l |= ((dim.lastColumn - dim.firstColumn) & (1 << 14) - 1) << 14;
            l |= 1;
        }
        sharedCalc[i] = new PreCalc(l);
    }

    /**
     * Setting calc string
     *
     * @param i the ref id
     * @param calc the calc string
     */
    void setCalc(int i, String calc) {
        if (sharedCalc == null || sharedCalc.length <= i
            || sharedCalc[i] == null || StringUtil.isEmpty(calc))
            return;

        sharedCalc[i].setCalc(calc.toCharArray());
    }

    /**
     * Get calc string by ref id and coordinate
     *
     * @param i the ref id
     * @param coordinate the cell coordinate
     * @return calc string
     */
    String getCalc(int i, long coordinate) {
        // Index out of range
        if (sharedCalc == null || sharedCalc.length <= i
            || sharedCalc[i] == null)
            return EMPTY;

        return sharedCalc[i].get(coordinate);
    }

    /**
     * Returns deep clone cells
     *
     * @return cells
     */
    public Cell[] copyCells() {
        return copyCells(cells.length);
    }

    /**
     * Returns deep clone cells
     *
     * @param newLength the length of the copy to be returned
     * @return cells
     */
    public Cell[] copyCells(int newLength) {
        Cell[] newCells = new Cell[newLength];
        int oldRow = cells.length;
        for (int k = 0; k < newLength; k++) {
            newCells[k] = new Cell((short) (k + 1));
            // Copy values
            if (k < oldRow && cells[k] != null) {
                newCells[k].from(cells[k]);
            }
        }
        return newCells;
    }

    /**
     * Setting custom {@link Cell}
     *
     * @param cells row cells
     * @return current Row
     */
    public Row setCells(Cell[] cells) {
        this.cells = cells;
        this.fc = 0;
        this.lc = cells.length;
        return this;
    }

    /**
     * Setting custom {@link Cell}
     *
     * @param cells row cells
     * @param fromIndex specify the first cells index(one base)
     * @param toIndex specify the last cells index(one base)
     * @return current Rows
     */
    public Row setCells(Cell[] cells, int fromIndex, int toIndex) {
        if (fromIndex < 0)
            throw new IndexOutOfBoundsException("fromIndex = " + fromIndex);
        if (toIndex > cells.length)
            throw new IndexOutOfBoundsException("toIndex = " + toIndex);
        if (fromIndex > toIndex)
            throw new IllegalArgumentException("fromIndex(" + fromIndex +
                ") > toIndex(" + toIndex + ")");

        this.cells = cells;
        this.fc = fromIndex;
        this.lc = toIndex;
        return this;
    }

    /**
     * Convert to column index
     *
     * @param cb character buffer
     * @param a the start index
     * @param b the end index
     * @return the cell index
     */
    public static int toCellIndex(char[] cb, int a, int b) {
        int n = 0;
        for (; a <= b; a++) {
            if (cb[a] <= 'Z' && cb[a] >= 'A') {
                n = n * 26 + cb[a] - '@';
            } else if (cb[a] <= 'z' && cb[a] >= 'a') {
                n = n * 26 + cb[a] - '、';
            } else break;
        }
        return n;
    }

    // -1: not a number
    // 0: empty
    // 1: int
    // 2: long
    // 3: double / decimal
    public static int testNumberType(char[] cb, int a, int b) {
        if (a == b) return 0;
        if (b - a == 1) return cb[a] >= '0' && cb[a] <= '9' ? 1 : -1;
        int dotIdx = -1, eIdx = -1, i = a, j;
        if (cb[i] == '-') i++;
        j = i;
        for ( ; i < b; ) {
            char c = cb[i++];
            if (c >= '0' && c <= '9') continue;
            else if (c == '.') {
                if (dotIdx >= 0 || eIdx >= 0) return -1;
                dotIdx = i - 1;
            }
            else if (c == 'e' || c == 'E') {
                if (eIdx > 0 || i == 1) return -1;
                eIdx = i - 1;
                if (i + 1 > b) return -1;
                c = cb[i++];
                if (c == '-' || c == '+') {
                    if (i + 1 > b) return -1;
                }
                else if (c < '0' || c > '9') return -1;
            }
            else return -1;
        }

//        int intPart = dotIdx == -1 ? eIdx == -1 ? b : eIdx : dotIdx, ePart = eIdx > 0 ? b - ep : 0, dotPart = dotIdx >= 0 ? (eIdx > 0 ? eIdx : b) - dotIdx - 1 : 0;

        if (b - j == 1 && dotIdx >= 0) return -1;
        return dotIdx >= 0 || eIdx > 1 ? 3 : b - j >= 10 ? 2 : 1;
    }
}

/**
 * Test and merge formula each rows.
 *
 * @author guanquan.wang at 2019-12-31 15:42
 */
@FunctionalInterface
interface MergeCalcFunc {

    /**
     * Merge formula in rows
     *
     * @param row thr row number
     * @param cells the cells in row
     * @param n count of cells
     */
    void accept(int row, Cell[] cells, int n);
}

/**
 * Test and copy value on merged cells
 *
 * @author guanquan.wang at 2020-01-17 11:36
 */
@FunctionalInterface
interface MergeValueFunc {

    /**
     * Copy merged values
     *
     * @param row thr row number
     * @param cell all cell in row
     */
    void accept(int row, Cell cell);
}
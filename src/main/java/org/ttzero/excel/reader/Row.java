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
 * 创建一个独立的对象而后者返回的对象是内存共享的，如果需要使用数组或集合类收集对象则需要使用{@code to}方法，{@code too}方法
 * 用于流式one-by-one的场景。
 *
 * <p>使用{@code to}和{@code too}方法转对象都有一个前提，那就是所转对象的属性或set方法必须使用{@link org.ttzero.excel.annotation.ExcelColumn}注解，
 * 通过表头行上的文本与&#x40;ExcelColumn注解的{@code value}值进行匹配，衣可以指定{@code colIndex}按照列索引进行匹配。
 * 如果使用{@link Sheet#forceImport}强制匹配时无&#x40;ExcelColumn注解的字段将会按照字段名进行匹配，</p>
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
    private PreCalc[] sharedCalc;

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
     * Returns the index of the last column (zero base, exclude).
     * The last index of column is increment the max available index
     *
     * @return the last column index
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
     * Test unused row (not contains any filled, formatted, border, value or other styles)
     *
     * @return true if unused
     */
    public boolean isEmpty() {
        return lc - fc <= 0;
    }

    /**
     * Returns {@code true} if any cell in row contains filled, formatted, border, value or other styles
     * otherwise returns {@code false}.
     *
     * This method exists to be used as a
     * {@link java.util.function.Predicate}, {@code filter(Row::nonEmpty)}
     *
     * @return {@code true} if any cell in row contains style and value
     * otherwise {@code false}
     *
     * @see java.util.function.Predicate
     * @see #isEmpty()
     */
    public boolean nonEmpty() {
        return lc > fc;
    }

    /**
     * Tests the value of all cells is null or whitespace
     *
     * @return true if all cell value is null
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
     * Returns {@code true} if any cell in row contains values
     * otherwise returns {@code false}.
     *
     * This method exists to be used as a
     * {@link java.util.function.Predicate}, {@code filter(Row::nonBlank)}
     *
     * @return {@code true} if any cell in row contains values
     * otherwise {@code false}
     *
     * @see java.util.function.Predicate
     * @see #isBlank()
     */
    public boolean nonBlank() {
        return !isBlank();
    }

    /**
     * 获取行高，只有{@code XMLFullRow}才会实际返回值
     *
     * @return 当{@code customHeight=1}时返回自定义行高否则返回{@code null}
     */
    public Double getHeight() {
        return null;
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
     * Returns {@link Cell}
     *
     * @param i the position of cell
     * @return the {@link Cell}
     */
    public Cell getCell(int i) {
        rangeCheck(i);
        return i < lc ? cells[i] : UNALLOCATED_CELL;
    }

    /**
     * Search {@link Cell} by column name
     *
     * @param name the column name
     * @return the {@link Cell}
     */
    public Cell getCell(String name) {
        int i = hr.getIndex(name);
        rangeCheck(i);
        return i < lc ? cells[i] : UNALLOCATED_CELL;
    }

    /**
     * convert row to header_row
     *
     * @return header Row
     */
    public HeaderRow asHeader() {
        return new HeaderRow().with(this);
    }

    /**
     * Setting header row
     *
     * @param hr {@link HeaderRow}
     * @return self
     */
    public Row setHr(HeaderRow hr) {
        this.hr = hr;
        return this;
    }

    /**
     * Get {@code Boolean} value by column index
     *
     * @param columnIndex the cell index
     * @return {@code Boolean}
     */
    public Boolean getBoolean(int columnIndex) {
        Cell c = getCell(columnIndex);
        return getBoolean(c);
    }

    /**
     * Get {@code Boolean} value by column name
     *
     * @param columnName the cell name
     * @return {@code Boolean}
     */
    public Boolean getBoolean(String columnName) {
        Cell c = getCell(columnName);
        return getBoolean(c);
    }

    /**
     * Get {@code Boolean} value
     *
     * @param c the {@link Cell}
     * @return {@code Boolean}
     */
    public Boolean getBoolean(Cell c) {
        boolean v;
        switch (c.t) {
            case BOOL       : v = c.bv;                                 break;
            case NUMERIC    : v = c.nv != 0;                            break;
            case LONG       : v = c.lv != 0L;                           break;
            case SST        : if (c.sv == null) c.setSv(sst.get(c.nv)); // @Mark:=>There is no missing `break`, this is normal logic here
            case INLINESTR  : v = "true".equalsIgnoreCase(c.sv);        break;
            case DECIMAL    : v = c.mv.compareTo(BigDecimal.ZERO) != 0; break;
            case DOUBLE     : v = c.dv != .0D;                          break;
            case BLANK      :
            case EMPTY_TAG  :
            case UNALLOCATED: return null;
            default         : v = false;
        }
        return v;
    }

    /**
     * Get {@code Byte} value by column index
     *
     * @param columnIndex the cell index
     * @return {@code Byte}
     */
    public Byte getByte(int columnIndex) {
        Cell c = getCell(columnIndex);
        return getByte(c);
    }

    /**
     * Get {@code Byte} value by column name
     *
     * @param columnName the cell name
     * @return {@code Byte}
     */
    public Byte getByte(String columnName) {
        Cell c = getCell(columnName);
        return getByte(c);
    }

    /**
     * Get {@code Byte} value
     *
     * @param c the {@link Cell}
     * @return {@code Byte}
     */
    public Byte getByte(Cell c) {
        byte b = 0;
        switch (c.t) {
            case NUMERIC    : b |= c.nv;                                break;
            case LONG       : b |= c.lv;                                break;
            case DECIMAL    : b = c.mv.byteValue();                     break;
            case DOUBLE     : b |= (int) c.dv;                          break;
            case BOOL       : b |= c.bv ? 1 : 0;                        break;
            default         : return null;
        }
        return b;
    }

    /**
     * Get {@code Character} value by column index
     *
     * @param columnIndex the cell index
     * @return {@code Character}
     */
    public Character getChar(int columnIndex) {
        Cell c = getCell(columnIndex);
        return getChar(c);
    }

    /**
     * Get {@code Character} value by column name
     *
     * @param columnName the cell name
     * @return {@code Character}
     */
    public Character getChar(String columnName) {
        Cell c = getCell(columnName);
        return getChar(c);
    }

    /**
     * Get {@code Character} value
     *
     * @param c the {@link Cell}
     * @return {@code Character}
     */
    public Character getChar(Cell c) {
        char cc = 0;
        switch (c.t) {
            case SST        : if (c.sv == null) c.setSv(sst.get(c.nv));   // @Mark:=>There is no missing `break`, this is normal logic here
            case INLINESTR  : if (isNotEmpty(c.sv)) cc = c.sv.charAt(0);  break;
            case NUMERIC    : cc |= c.nv;                                 break;
            case LONG       : cc |= c.lv;                                 break;
            case BOOL       : cc |= c.bv ? 1 : 0;                         break;
            case DECIMAL    : cc |= c.mv.intValue();                      break;
            case DOUBLE     : cc |= (int) c.dv;                           break;
            default         : return null;
        }
        return cc;
    }

    /**
     * Get {@code Short} value by column index
     *
     * @param columnIndex the cell index
     * @return {@code Short}
     */
    public Short getShort(int columnIndex) {
        Cell c = getCell(columnIndex);
        return getShort(c);
    }

    /**
     * Get {@code Short} value by column name
     *
     * @param columnName the cell name
     * @return {@code Short}
     */
    public Short getShort(String columnName) {
        Cell c = getCell(columnName);
        return getShort(c);
    }

    /**
     * Get {@code Short} value
     *
     * @param c the {@link Cell}
     * @return {@code Short}
     */
    public Short getShort(Cell c) {
        short s = 0;
        switch (c.t) {
            case NUMERIC    : s |= c.nv;                                break;
            case LONG       : s |= c.lv;                                break;
            case DECIMAL    : s = c.mv.shortValue();                    break;
            case DOUBLE     : s |= (int) c.dv;                          break;
            case SST        : if (c.sv == null) c.setSv(sst.get(c.nv)); // @Mark:=>There is no missing `break`, this is normal logic here
            case INLINESTR  :
                if (StringUtil.isEmpty(c.sv)) return null;
                String ss = c.sv.trim();
                int t = testNumberType(ss.toCharArray(), 0, ss.length());
                switch (t) {
                    case 1  : s |= Integer.parseInt(ss);                break;
                    case 2  : s |= Long.parseLong(ss);                  break;
                    case 3  : s = (short) Double.parseDouble(ss);       break;
                    case 0  : return null;
                    default : throw new NumberFormatException("For input string: \"" + c.sv + "\"");
                }                                                       break;
            case BOOL       : s |= c.bv ? 1 : 0;                        break;
            default         : return null;
        }
        return s;
    }

    /**
     * Get {@code Integer} value by column index
     *
     * @param columnIndex the cell index
     * @return {@code Integer}
     */
    public Integer getInt(int columnIndex) {
        Cell c = getCell(columnIndex);
        return getInt(c);
    }

    /**
     * Get {@code Integer} value by column name
     *
     * @param columnName the cell name
     * @return {@code Integer}
     */
    public Integer getInt(String columnName) {
        Cell c = getCell(columnName);
        return getInt(c);
    }

    /**
     * Get {@code Integer} value
     *
     * @param c the {@link Cell}
     * @return {@code Integer}
     */
    public Integer getInt(Cell c) {
        int n = 0;
        switch (c.t) {
            case NUMERIC    : n = c.nv;                                 break;
            case LONG       : n = (int) c.lv;                           break;
            case DECIMAL    : n = c.mv.intValue();                      break;
            case DOUBLE     : n = (int) c.dv;                           break;
            case SST        : if (c.sv == null) c.setSv(sst.get(c.nv)); // @Mark:=>There is no missing `break`, this is normal logic here
            case INLINESTR  :
                if (StringUtil.isEmpty(c.sv)) return null;
                String ss = c.sv.trim();
                int t = testNumberType(ss.toCharArray(), 0, ss.length());
                switch (t) {
                    case 1  : n = Integer.parseInt(ss);                 break;
                    case 2  : n |= Long.parseLong(ss);                  break;
                    case 3  : n = (int) Double.parseDouble(ss);         break;
                    case 0  : return null;
                    default : throw new NumberFormatException("For input string: \"" + c.sv + "\"");
                }                                                       break;
            case BOOL       : n = c.bv ? 1 : 0;                         break;
            default         : return null;
        }
        return n;
    }

    /**
     * Get {@code Long} value by column index
     *
     * @param columnIndex the cell index
     * @return {@code Long}
     */
    public Long getLong(int columnIndex) {
        Cell c = getCell(columnIndex);
        return getLong(c);
    }

    /**
     * Get {@code Long} value by column name
     *
     * @param columnName the cell name
     * @return {@code Long}
     */
    public Long getLong(String columnName) {
        Cell c = getCell(columnName);
        return getLong(c);
    }

    /**
     * Get {@code Long} value
     *
     * @param c the {@link Cell}
     * @return {@code Long}
     */
    public Long getLong(Cell c) {
        long l;
        switch (c.t) {
            case LONG       : l = c.lv;                                 break;
            case NUMERIC    : l = c.nv;                                 break;
            case DECIMAL    : l = c.mv.longValue();                     break;
            case DOUBLE     : l = (long) c.dv;                          break;
            case SST        : if (c.sv == null) c.setSv(sst.get(c.nv)); // @Mark:=>There is no missing `break`, this is normal logic here
            case INLINESTR  :
                if (StringUtil.isEmpty(c.sv)) return null;
                String ss = c.sv.trim();
                int t = testNumberType(ss.toCharArray(), 0, ss.length());
                switch (t) {
                    case 1  :
                    case 2  : l = Long.parseLong(ss);                   break;
                    case 3  : l = (long) Double.parseDouble(ss);        break;
                    case 0  : return null;
                    default : throw new NumberFormatException("For input string: \"" + c.sv + "\"");
                }                                                       break;
            case BOOL       : l = c.bv ? 1L : 0L;                       break;
            default         : return null;
        }
        return l;
    }

    /**
     * Get string value by column index
     *
     * @param columnIndex the cell index
     * @return string
     */
    public String getString(int columnIndex) {
        Cell c = getCell(columnIndex);
        return getString(c);
    }

    /**
     * Get string value by column name
     *
     * @param columnName the cell name
     * @return string
     */
    public String getString(String columnName) {
        Cell c = getCell(columnName);
        return getString(c);
    }

    /**
     * Get string value
     *
     * @param c the {@link Cell}
     * @return string
     */
    public String getString(Cell c) {
        String s;
        switch (c.t) {
            case SST        : if (c.sv == null) c.setSv(sst.get(c.nv)); // @Mark:=>There is no missing `break`, this is normal logic here
            case INLINESTR  : s = c.sv;                                 break;
            case BLANK      :
            case EMPTY_TAG  :
            case UNALLOCATED: s = null;                                 break;
            case LONG       : s = String.valueOf(c.lv);                 break;
            case NUMERIC    : s = String.valueOf(c.nv);                 break;
            case DECIMAL    : s = c.mv.toString();                      break;
            case DOUBLE     : s = String.valueOf(c.dv);                 break;
            case BOOL       : s = c.bv ? "true" : "false";              break;
            default         : s = c.sv;
        }
        return s;
    }

    /**
     * Get {@code Float} value by column index
     *
     * @param columnIndex the cell index
     * @return {@code Float}
     */
    public Float getFloat(int columnIndex) {
        Double d = getDouble(columnIndex);
        return d != null ? Float.valueOf(d.toString()) : null;
    }

    /**
     * Get {@code Float} value by column index
     *
     * @param columnName the cell index
     * @return {@code Float}
     */
    public Float getFloat(String columnName) {
        Double d = getDouble(columnName);
        return d != null ? Float.valueOf(d.toString()) : null;
    }

    /**
     * Get {@code Float} value by cell
     *
     * @param c the {@link Cell}
     * @return {@code Float}
     */
    public Float getFloat(Cell c) {
        Double d = getDouble(c);
        return d != null ? Float.valueOf(d.toString()) : null;
    }

    /**
     * Get {@code Double} value by column index
     *
     * @param columnIndex the cell index
     * @return {@code Double}
     */
    public Double getDouble(int columnIndex) {
        Cell c = getCell(columnIndex);
        return getDouble(c);
    }

    /**
     * Get {@code Double} value by column name
     *
     * @param columnName the cell name
     * @return {@code Double}
     */
    public Double getDouble(String columnName) {
        Cell c = getCell(columnName);
        return getDouble(c);
    }

    /**
     * Get {@code Double} value
     *
     * @param c the {@link Cell}
     * @return {@code Double}
     */
    public Double getDouble(Cell c) {
        double d;
        switch (c.t) {
            case DECIMAL    : d = c.mv.doubleValue();                   break;
            case DOUBLE     : d = c.dv;                                 break;
            case NUMERIC    : d = c.nv;                                 break;
            case LONG       : d = c.lv;                                 break;
            case SST        : if (c.sv == null) c.setSv(sst.get(c.nv)); // @Mark:=>There is no missing `break`, this is normal logic here
            case INLINESTR  :
                if (isNotBlank(c.sv)) d = Double.parseDouble(c.sv.trim());
                else return null;                                       break;
            default         : return null;
        }
        return d;
    }

    /**
     * Get {@link java.math.BigDecimal} value by column index
     *
     * @param columnIndex the cell index
     * @return BigDecimal
     */
    public BigDecimal getDecimal(int columnIndex) {
        Cell c = getCell(columnIndex);
        return getDecimal(c);
    }

    /**
     * Get {@link java.math.BigDecimal} value by column name
     *
     * @param columnName the cell name
     * @return BigDecimal
     */
    public BigDecimal getDecimal(String columnName) {
        Cell c = getCell(columnName);
        return getDecimal(c);
    }

    /**
     * Get {@link java.math.BigDecimal} value
     *
     * @param c the {@link Cell}
     * @return BigDecimal
     */
    public BigDecimal getDecimal(Cell c) {
        BigDecimal bd;
        switch (c.t) {
            case DECIMAL    : bd = c.mv;                                break;
            case DOUBLE     : bd = BigDecimal.valueOf(c.dv);            break;
            case NUMERIC    : bd = BigDecimal.valueOf(c.nv);            break;
            case LONG       : bd = BigDecimal.valueOf(c.lv);            break;
            case SST        : if (c.sv == null) c.setSv(sst.get(c.nv)); // @Mark:=>There is no missing `break`, this is normal logic here
            case INLINESTR  : bd = isNotBlank(c.sv) ? new BigDecimal(c.sv.trim()) : null; break;
            default         : bd = null;
        }
        return bd;
    }

    /**
     * Get {@link java.util.Date} value by column index
     *
     * @param columnIndex the cell index
     * @return Date
     */
    public Date getDate(int columnIndex) {
        Cell c = getCell(columnIndex);
        return getDate(c);
    }

    /**
     * Get {@link java.util.Date} value by column name
     *
     * @param columnName the cell name
     * @return Date
     */
    public Date getDate(String columnName) {
        Cell c = getCell(columnName);
        return getDate(c);
    }

    /**
     * Get {@link java.util.Date} value
     *
     * @param c the {@link Cell}
     * @return BigDecimal
     */
    public Date getDate(Cell c) {
        Date date;
        switch (c.t) {
            case NUMERIC    : date = toDate(c.nv);                      break;
            case DECIMAL    : date = toDate(c.mv.doubleValue());        break;
            case DOUBLE     : date = toDate(c.dv);                      break;
            case SST        : if (c.sv == null) c.setSv(sst.get(c.nv)); // @Mark:=>There is no missing `break`, this is normal logic here
            case INLINESTR  : date = isNotBlank(c.sv) ? toDate(c.sv.trim()) : null; break;
            default         : date = null;
        }
        return date;
    }

    /**
     * Get {@link java.sql.Timestamp} value by column index
     *
     * @param columnIndex the cell index
     * @return java.sql.Timestamp
     */
    public Timestamp getTimestamp(int columnIndex) {
        Cell c = getCell(columnIndex);
        return getTimestamp(c);
    }

    /**
     * Get {@link java.sql.Timestamp} value by column name
     *
     * @param columnName the cell name
     * @return java.sql.Timestamp
     */
    public Timestamp getTimestamp(String columnName) {
        Cell c = getCell(columnName);
        return getTimestamp(c);
    }

    /**
     * Get {@link java.sql.Timestamp} value
     *
     * @param c the {@link Cell}
     * @return java.sql.Timestamp
     */
    public Timestamp getTimestamp(Cell c) {
        Timestamp ts;
        switch (c.t) {
            case NUMERIC    : ts = toTimestamp(c.nv);                   break;
            case DECIMAL    : ts = toTimestamp(c.mv.doubleValue());     break;
            case DOUBLE     : ts = toTimestamp(c.dv);                   break;
            case SST        : if (c.sv == null) c.setSv(sst.get(c.nv)); // @Mark:=>There is no missing `break`, this is normal logic here
            case INLINESTR  : ts = isNotBlank(c.sv) ? toTimestamp(c.sv.trim()) : null; break;
            default         : ts = null;
        }
        return ts;
    }

    /**
     * Get {@link java.sql.Time} value by column index
     *
     * @param columnIndex the cell index
     * @return java.sql.Time
     */
    public java.sql.Time getTime(int columnIndex) {
        return getTime(getCell(columnIndex));
    }

    /**
     * Get {@link java.sql.Time} value by column name
     *
     * @param columnName the cell name
     * @return java.sql.Time
     */
    public java.sql.Time getTime(String columnName) {
        return getTime(getCell(columnName));
    }

    /**
     * Get {@link java.sql.Time} value by column name
     *
     * @param c the {@link Cell}
     * @return java.sql.Time
     */
    public java.sql.Time getTime(Cell c) {
        java.sql.Time t;
        switch (c.t) {
            case DECIMAL    : t = toTime(c.mv.doubleValue());           break;
            case DOUBLE     : t = toTime(c.dv);                         break;
            case SST        : if (c.sv == null) c.setSv(sst.get(c.nv)); // @Mark:=>There is no missing `break`, this is normal logic here
            case INLINESTR  : t = isNotBlank(c.sv) ? toTime(c.sv.trim()) : null; break;
            default         : t = null;
        }
        return t;
    }

    /**
     * Get {@link LocalDateTime} value by column index
     *
     * @param columnIndex the cell index
     * @return java.time.LocalDateTime
     */
    public LocalDateTime getLocalDateTime(int columnIndex) {
        Cell c = getCell(columnIndex);
        return getLocalDateTime(c);
    }

    /**
     * Get {@link LocalDateTime} value by column name
     *
     * @param columnName the cell name
     * @return java.time.LocalDateTime
     */
    public LocalDateTime getLocalDateTime(String columnName) {
        Cell c = getCell(columnName);
        return getLocalDateTime(c);
    }

    /**
     * Get {@link LocalDateTime} value
     *
     * @param c the {@link Cell}
     * @return java.time.LocalDateTime
     */
    public LocalDateTime getLocalDateTime(Cell c) {
        LocalDateTime ldt;
        switch (c.t) {
            case NUMERIC    : ldt = toLocalDateTime(c.nv);              break;
            case DECIMAL    : ldt = toLocalDateTime(c.mv.doubleValue());break;
            case DOUBLE     : ldt = toLocalDateTime(c.dv);              break;
            case SST        : if (c.sv == null) c.setSv(sst.get(c.nv)); // @Mark:=>There is no missing `break`, this is normal logic here
            case INLINESTR  : ldt = isNotBlank(c.sv) ? toTimestamp(c.sv.trim()).toLocalDateTime() : null; break;
            default         : ldt = null;
        }
        return ldt;
    }

    /**
     * Get {@link LocalDate} value by column index
     *
     * @param columnIndex the cell index
     * @return java.time.LocalDate
     */
    public LocalDate getLocalDate(int columnIndex) {
        Cell c = getCell(columnIndex);
        return getLocalDate(c);
    }

    /**
     * Get {@link LocalDate} value by column name
     *
     * @param columnName the cell name
     * @return java.time.LocalDate
     */
    public LocalDate getLocalDate(String columnName) {
        Cell c = getCell(columnName);
        return getLocalDate(c);
    }

    /**
     * Get {@link LocalDate} value
     *
     * @param c the {@link Cell}
     * @return java.time.LocalDate
     */
    public LocalDate getLocalDate(Cell c) {
        LocalDate ld;
        switch (c.t) {
            case NUMERIC    : ld = toLocalDate(c.nv);                   break;
            case DECIMAL    : ld = toLocalDate(c.mv.intValue());        break;
            case DOUBLE     : ld = toLocalDate((int) c.dv);             break;
            case SST        : if (c.sv == null) c.setSv(sst.get(c.nv)); // @Mark:=>There is no missing `break`, this is normal logic here
            case INLINESTR  : ld = isNotBlank(c.sv) ? toTimestamp(c.sv.trim()).toLocalDateTime().toLocalDate() : null; break;
            default         : ld = null;
        }
        return ld;
    }

    /**
     * Get {@link LocalTime} value by column index
     *
     * @param columnIndex the cell index
     * @return java.time.LocalTime
     */
    public LocalTime getLocalTime(int columnIndex) {
        Cell c = getCell(columnIndex);
        return getLocalTime(c);
    }

    /**
     * Get {@link LocalTime} value by column name
     *
     * @param columnName the cell name
     * @return java.time.LocalTime
     */
    public LocalTime getLocalTime(String columnName) {
        Cell c = getCell(columnName);
        return getLocalTime(c);
    }

    /**
     * Get {@link LocalTime} value
     *
     * @param c the {@link Cell}
     * @return java.time.LocalTime
     */
    public LocalTime getLocalTime(Cell c) {
        LocalTime lt;
        switch (c.t) {
            case NUMERIC     : lt = toLocalTime(c.nv);                  break;
            case DECIMAL     : lt = toLocalTime(c.mv.doubleValue());    break;
            case DOUBLE      : lt = toLocalTime(c.dv);                  break;
            case SST         : if (c.sv == null) c.setSv(sst.get(c.nv));// @Mark:=>There is no missing `break`, this is normal logic here
            case INLINESTR   :
                if (isNotBlank(c.sv)) {
                    c.sv = c.sv.trim();
                    // 00:00:00
                    if (c.sv.length() == 8 && c.sv.charAt(2) == ':' && c.sv.charAt(5) == ':') lt = toLocalTime(c.sv);
                    else lt = toTimestamp(c.sv).toLocalDateTime().toLocalTime();
                } else lt = null;
                break;
            default          : lt = null;
        }
        return lt;
    }

    /**
     * Returns formula if exists
     *
     * @param columnIndex the cell index
     * @return the formula string if exists, otherwise return null
     */
    public String getFormula(int columnIndex) {
        Cell c = getCell(columnIndex);
        return c.fv;
    }

    /**
     * Returns formula if exists
     *
     * @param columnName the cell name
     * @return the formula string if exists, otherwise return null
     */
    public String getFormula(String columnName) {
        Cell c = getCell(columnName);
        return c.fv;
    }

    /**
     * Check cell has formula
     *
     * @param columnIndex the cell index
     * @return the formula string if exists, otherwise return null
     */
    public boolean hasFormula(int columnIndex) {
        return getCell(columnIndex).f;
    }

    /**
     * Check cell has formula
     *
     * @param columnName the cell name
     * @return the formula string if exists, otherwise return null
     */
    public boolean hasFormula(String columnName) {
        return getCell(columnName).f;
    }

    /**
     * Returns the type of cell
     *
     * @param columnIndex the cell index from zero
     * @return the {@link CellType}
     */
    public CellType getCellType(int columnIndex) {
        Cell c = getCell(columnIndex);
        return getCellType(c);
    }

    /**
     * Returns the type of cell
     *
     * @param columnName the cell name
     * @return the {@link CellType}
     */
    public CellType getCellType(String columnName) {
        Cell c = getCell(columnName);
        return getCellType(c);
    }

    /**
     * Returns the type of cell
     *
     * @param c the {@link Cell}
     * @return the {@link CellType}
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
     * Returns the cell styles
     *
     * @param columnIndex the cell index from zero
     * @return the style value
     */
    public int getCellStyle(int columnIndex) {
        Cell c = getCell(columnIndex);
        return getCellStyle(c);
    }

    /**
     * Returns the cell styles
     *
     * @param columnName the cell name
     * @return the style value
     */
    public int getCellStyle(String columnName) {
        Cell c = getCell(columnName);
        return getCellStyle(c);
    }

    /**
     * Returns the cell styles
     *
     * @param c the {@link Cell}
     * @return the style value
     */
    public int getCellStyle(Cell c) {
        return styles.getStyleByIndex(c.xf);
    }

    /**
     * Tests the specify cell value is blank
     *
     * @param columnIndex the cell index
     * @return true if cell value is blank
     */
    public boolean isBlank(int columnIndex) {
        Cell c = getCell(columnIndex);
        return isBlank(c);
    }

    /**
     * Tests the specify cell value is blank
     *
     * @param columnName the cell name
     * @return true if cell value is blank
     */
    public boolean isBlank(String columnName) {
        Cell c = getCell(columnName);
        return isBlank(c);
    }

    /**
     * Tests the specify cell value is blank
     *
     * @param c the {@link Cell}
     * @return true if cell value is blank
     */
    public boolean isBlank(Cell c) {
        boolean blank;
        switch (c.t) {
            case SST        : if (c.sv == null) c.setSv(sst.get(c.nv)); // @Mark:=>There is no missing `break`, this is normal logic here
            case INLINESTR  : blank = StringUtil.isBlank(c.sv); break;
            case BLANK      :
            case EMPTY_TAG  :
            case UNALLOCATED: blank = true;                     break;
            default         : blank = false;
        }
        return blank;
    }

    /**
     * Returns the binding type if is bound, otherwise returns Row
     *
     * @param <T> the type of binding
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
     * Returns the binding type if is bound, otherwise returns Row
     *
     * @param <T> the type of binding
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
     * @param clazz the type of binding
     * @param <T>   the type of return object
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
     * Convert to T object, support annotation
     * the is a memory shared object
     *
     * @param clazz the type of binding
     * @param <T>   the type of return object
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
                case SST      : if (c.sv == null) c.setSv(sst.get(c.nv)); // @Mark:=>There is no missing `break`, this is normal logic here
                case INLINESTR: joiner.add(c.sv);                 break;
                case NUMERIC  :
                    if (!styles.fastTestDateFmt(c.xf)) joiner.add(String.valueOf(c.nv));
                    else joiner.add(toLocalDate(c.nv).toString());
                    break;
                case LONG     : joiner.add(String.valueOf(c.lv)); break;
                case DECIMAL:
                    if (!styles.fastTestDateFmt(c.xf)) joiner.add(c.mv.toString());
                    else if (c.mv.compareTo(BigDecimal.ONE) > 0) joiner.add(toTimestamp(c.mv.doubleValue()).toString());
                    else joiner.add(toLocalTime(c.mv.doubleValue()).toString());
                    break;
                case DOUBLE:
                    if (!styles.fastTestDateFmt(c.xf)) joiner.add(String.valueOf(c.dv));
                    else if (c.dv > 1.0000) joiner.add(toTimestamp(c.dv).toString());
                    else joiner.add(toLocalTime(c.dv).toString());
                    break;
                case BLANK    :
                case EMPTY_TAG: joiner.add(EMPTY);                break;
                case BOOL     : joiner.add(String.valueOf(c.bv)); break;
                default       : joiner.add(null);
            }
        }
        return joiner.toString();
    }

    /**
     * Convert row data to LinkedMap(sort by column index)
     *
     * @return the key is name or index(if not name here)
     */
    public Map<String, Object> toMap() {
        if (isEmpty() || hr == null) return Collections.emptyMap();
        // Maintain the column orders
        Map<String, Object> data = new LinkedHashMap<>(Math.max(16, hr.lc - hr.fc));
        String[] names = hr.names;
        String key;
        for (int i = hr.fc; i < hr.lc; i++) {
            Cell c = cells[i];
            key = names[i];
            // Ignore null key
            if (key == null) continue;
            switch (c.t) {
                case SST      : if (c.sv == null) c.setSv(sst.get(c.nv)); // @Mark:=>There is no missing `break`, this is normal logic here
                case INLINESTR: data.put(key, c.sv);              break;
                case NUMERIC  :
                    if (!styles.fastTestDateFmt(c.xf)) data.put(key, c.nv);
                    else data.put(key, toTimestamp(c.nv));
                    break;
                case LONG     :  data.put(key, c.lv);             break;
                case DECIMAL  :
                    if (!styles.fastTestDateFmt(c.xf)) data.put(key, c.mv);
                    else if (c.mv.compareTo(BigDecimal.ONE) > 0) data.put(key, toTimestamp(c.mv.doubleValue()));
                    else data.put(key, toTime(c.mv.doubleValue()));
                    break;
                case DOUBLE   :
                    if (!styles.fastTestDateFmt(c.xf)) data.put(key, c.dv);
                    else if (c.dv > 1.00000) data.put(key, toTimestamp(c.dv));
                    else data.put(key, toTime(c.dv));
                    break;
                case BLANK    :
                case EMPTY_TAG: data.put(key, EMPTY);             break;
                case BOOL     : data.put(key, c.bv);              break;
                default       : data.put(key, null);
            }
        }
        return data;
    }

    /**
     * Add function shared ref
     * <blockquote><pre>
     * 63   : Not used
     * 42-62: First row number
     * 28-41: First column number
     * 8-27/14-27: Size, if axis is zero the size used 20 bits, otherwise used 14 bits
     * 2-7/2-13: Not used
     * 0-1    : Axis, 00: range 01: y-axis 10: x-axis
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
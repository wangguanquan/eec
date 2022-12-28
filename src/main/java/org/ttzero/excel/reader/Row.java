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
 * @author guanquan.wang at 2019-04-17 11:08
 */
public abstract class Row {
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
     * The number of row. (one base)
     *
     * @return int value
     * @deprecated replace with {@link #getRowNum()}
     */
    @Deprecated
    public int getRowNumber() {
        return index;
    }

    /**
     * The number of row. (one base)
     *
     * @return int value
     */
    public int getRowNum() {
        return index;
    }

    /**
     * Returns the index of the first column (zero base)
     *
     * @return the first column index
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
     * Test unused row (not contains any filled or formatted or value)
     *
     * @return true if unused
     */
    public boolean isEmpty() {
        return lc - fc <= 0;
    }

    /**
     * Tests the value of all cells is null or whitespace
     *
     * @return true if all cell value is null
     */
    public boolean isBlank() {
        if (lc > fc) {
            boolean blank;
            for (int i = fc; i < lc; i++) {
                Cell c = cells[i];
                switch (c.t) {
                    case SST:
                        if (c.sv == null) {
                            c.setSv(sst.get(c.nv));
                        }
                        // @Mark:=>There is no missing `break`, this is normal logic here
                    case INLINESTR:
                        blank = StringUtil.isBlank(c.sv);
                        break;
                    case BLANK:
                    case EMPTY_TAG:
                    case UNALLOCATED:
                        blank = true;
                        break;
                    default: blank = false;
                }
                if (!blank) return false;
            }
        }
        return true;
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
    protected HeaderRow asHeader() {
        return new HeaderRow().with(this);
    }

    /**
     * Setting header row
     *
     * @param hr {@link HeaderRow}
     * @return self
     */
    protected Row setHr(HeaderRow hr) {
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
            case BOOL:
                v = c.bv;
                break;
            case NUMERIC:
            case DOUBLE:
                v = c.nv != 0 || c.dv >= 0.000001 || c.dv <= -0.000001;
                break;
            case SST:
                if (c.sv == null) {
                    c.setSv(sst.get(c.nv));
                }
            // @Mark:=>There is no missing `break`, this is normal logic here
            case INLINESTR:
                v = c.sv != null && isNotBlank(c.sv.trim());
                break;
            case BLANK:
            case EMPTY_TAG:
            case UNALLOCATED:
                return null;
            default: v = false;
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
            case NUMERIC:
                b |= c.nv;
                break;
            case LONG:
                b |= c.lv;
                break;
            case BOOL:
                b |= c.bv ? 1 : 0;
                break;
            case DOUBLE:
                b |= (int) c.dv;
                break;
//            case BLANK:
//            case EMPTY_TAG:
//            case UNALLOCATED:
            default:
                return null;
//            default: throw new UncheckedTypeException("Can't convert cell value to Byte");
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
            case SST:
                if (c.sv == null) {
                    c.setSv(sst.get(c.nv));
                }
            // @Mark:=>There is no missing `break`, this is normal logic here
            case INLINESTR:
                if (isNotEmpty(c.sv)) {
                    cc |= c.sv.charAt(0);
                }
                break;
            case NUMERIC:
                cc |= c.nv;
                break;
            case LONG:
                cc |= c.lv;
                break;
            case BOOL:
                cc |= c.bv ? 1 : 0;
                break;
            case DOUBLE:
                cc |= (int) c.dv;
                break;
//            case BLANK:
//            case EMPTY_TAG:
//            case UNALLOCATED:
            default:
                return null;
//            default: throw new UncheckedTypeException("Can't convert cell value to Character");
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
            case NUMERIC:
                s |= c.nv;
                break;
            case LONG:
                s |= c.lv;
                break;
            case DOUBLE:
                s |= (int) c.dv;
                break;
            case SST:
                if (c.sv == null) {
                    c.setSv(sst.get(c.nv));
                }
            // @Mark:=>There is no missing `break`, this is normal logic here
            case INLINESTR:
                if (isNotBlank(c.sv)) {
                    c.sv = c.sv.trim();
                    if (c.sv.indexOf('E') >= 0 || c.sv.indexOf('e') >= 0) {
                        s = (short) Double.parseDouble(c.sv);
                    } else {
                        s = Long.valueOf(c.sv).shortValue();
                    }
                } else return null;
                break;
            case BOOL:
                s |= c.bv ? 1 : 0;
                break;
//            case BLANK:
//            case EMPTY_TAG:
//            case UNALLOCATED:
            default:
                return null;
//            default: throw new UncheckedTypeException("Can't convert cell value to short");
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
        int n;
        switch (c.t) {
            case NUMERIC:
                n = c.nv;
                break;
            case LONG:
                n = (int) c.lv;
                break;
            case DOUBLE:
                n = (int) c.dv;
                break;
            case SST:
                if (c.sv == null) {
                    c.setSv(sst.get(c.nv));
                }
            // @Mark:=>There is no missing `break`, this is normal logic here
            case INLINESTR:
                if (isNotBlank(c.sv)) {
                    c.sv = c.sv.trim();
                    if (c.sv.indexOf('E') >= 0 || c.sv.indexOf('e') >= 0) {
                        n = (int) Double.parseDouble(c.sv);
                    } else {
                        n = Long.valueOf(c.sv).intValue();
                    }
                } else return null;
                break;
            case BOOL:
                n = c.bv ? 1 : 0;
                break;
//            case BLANK:
//            case EMPTY_TAG:
//            case UNALLOCATED:
            default:
                return null;

//            default: throw new UncheckedTypeException("Can't convert cell value to Integer");
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
            case LONG:
                l = c.lv;
                break;
            case NUMERIC:
                l = c.nv;
                break;
            case DOUBLE:
                l = (long) c.dv;
                break;
            case SST:
                if (c.sv == null) {
                    c.setSv(sst.get(c.nv));
                }
            // @Mark:=>There is no missing `break`, this is normal logic here
            case INLINESTR:
                if (isNotBlank(c.sv)) {
                    c.sv = c.sv.trim();
                    if (c.sv.indexOf('E') >= 0 || c.sv.indexOf('e') >= 0) {
                        l = (long) Double.parseDouble(c.sv);
                    } else {
                        l = Long.parseLong(c.sv);
                    }
                } else return null;
                break;
            case BOOL:
                l = c.bv ? 1L : 0L;
                break;
//            case BLANK:
//            case EMPTY_TAG:
//            case UNALLOCATED:
            default:
                return null;
//            default: throw new UncheckedTypeException("Can't convert cell value to long");
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
            case SST:
                if (c.sv == null) {
                    c.setSv(sst.get(c.nv));
                }
                // @Mark:=>There is no missing `break`, this is normal logic here
            case INLINESTR:
                s = c.sv;
                break;
            case BLANK:
            case EMPTY_TAG:
            case UNALLOCATED:
                s = null;
                break;
            case LONG:
                s = String.valueOf(c.lv);
                break;
            case NUMERIC:
                s = String.valueOf(c.nv);
                break;
            case DOUBLE:
                s = String.valueOf(c.dv);
                break;
            case BOOL:
                s = c.bv ? "true" : "false";
                break;
            default: s = c.sv;
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
            case DOUBLE:
                d = c.dv;
                break;
            case NUMERIC:
                d = c.nv;
                break;
            case LONG:
                d = c.lv;
                break;
            case SST:
                if (c.sv == null) {
                    c.setSv(sst.get(c.nv));
                }
            // @Mark:=>There is no missing `break`, this is normal logic here
            case INLINESTR:
                if (isNotBlank(c.sv)) {
                    d = Double.parseDouble(c.sv.trim());
                } else return null;
                break;
//            case BLANK:
//            case EMPTY_TAG:
//            case UNALLOCATED:
            default:
                return null;
//            default: throw new UncheckedTypeException("Can't convert cell value to double");
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
            case DOUBLE:
                bd = BigDecimal.valueOf(c.dv);
                break;
            case NUMERIC:
                bd = BigDecimal.valueOf(c.nv);
                break;
            case LONG:
                bd = BigDecimal.valueOf(c.lv);
                break;
            case SST:
                if (c.sv == null) {
                    c.setSv(sst.get(c.nv));
                }
                // @Mark:=>There is no missing `break`, this is normal logic here
            case INLINESTR:
                bd = isNotBlank(c.sv) ? new BigDecimal(c.sv.trim()) : null;
                break;
//            case UNALLOCATED:
//            case BLANK:
//            case EMPTY_TAG:
            default:
                bd = null;
//                break;
//            default: throw new UncheckedTypeException("Can't convert cell value to java.math.BigDecimal");
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
            case NUMERIC:
                date = toDate(c.nv);
                break;
            case DOUBLE:
                date = toDate(c.dv);
                break;
            case SST:
                if (c.sv == null) {
                    c.setSv(sst.get(c.nv));
                }
                // @Mark:=>There is no missing `break`, this is normal logic here
            case INLINESTR:
                date = isNotBlank(c.sv) ? toDate(c.sv.trim()) : null;
                break;
//            case UNALLOCATED:
//            case BLANK:
//            case EMPTY_TAG:
            default:
                date = null;
//                break;
//            default: throw new UncheckedTypeException("Can't convert cell value to java.util.Date");
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
            case NUMERIC:
                ts = toTimestamp(c.nv);
                break;
            case DOUBLE:
                ts = toTimestamp(c.dv);
                break;
            case SST:
                if (c.sv == null) {
                    c.setSv(sst.get(c.nv));
                }
                // @Mark:=>There is no missing `break`, this is normal logic here
            case INLINESTR:
                ts = isNotBlank(c.sv) ? toTimestamp(c.sv.trim()) : null;
                break;
//            case UNALLOCATED:
//            case BLANK:
//            case EMPTY_TAG:
            default:
                ts = null;
//                break;
//            default: throw new UncheckedTypeException("Can't convert cell value to java.sql.Timestamp");
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
            case DOUBLE: t = toTime(c.dv); break;
            case SST:
                if (c.sv == null) {
                    c.setSv(sst.get(c.nv));
                }
                // @Mark:=>There is no missing `break`, this is normal logic here
            case INLINESTR: t = isNotBlank(c.sv) ? toTime(c.sv.trim()) : null; break;
//            case UNALLOCATED:
//            case BLANK:
//            case EMPTY_TAG:
            default:
                t = null;
//                break;
//            default:
//                throw new UncheckedTypeException("Can't convert cell value to java.sql.Time");
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
            case NUMERIC:
                ldt = toLocalDateTime(c.nv);
                break;
            case DOUBLE:
                ldt = toLocalDateTime(c.dv);
                break;
            case SST:
                if (c.sv == null) {
                    c.setSv(sst.get(c.nv));
                }
                // @Mark:=>There is no missing `break`, this is normal logic here
            case INLINESTR:
                ldt = isNotBlank(c.sv) ? toTimestamp(c.sv.trim()).toLocalDateTime() : null;
                break;
//            case UNALLOCATED:
//            case BLANK:
//            case EMPTY_TAG:
            default:
                ldt = null;
//                break;
//            default: throw new UncheckedTypeException("Can't convert cell value to java.time.LocalDateTime");
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
            case NUMERIC:
                ld = toLocalDate(c.nv);
                break;
            case DOUBLE:
                ld = toLocalDate((int) c.dv);
                break;
            case SST:
                if (c.sv == null) {
                    c.setSv(sst.get(c.nv));
                }
                // @Mark:=>There is no missing `break`, this is normal logic here
            case INLINESTR:
                ld = isNotBlank(c.sv) ? toTimestamp(c.sv.trim()).toLocalDateTime().toLocalDate() : null;
                break;
//            case UNALLOCATED:
//            case BLANK:
//            case EMPTY_TAG:
            default:
                ld = null;
//                break;
//            default: throw new UncheckedTypeException("Can't convert cell value to java.sql.Timestamp");
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
            case NUMERIC:
                lt = toLocalTime(c.nv);
                break;
            case DOUBLE:
                lt = toLocalTime(c.dv);
                break;
            case SST:
                if (c.sv == null) {
                    c.setSv(sst.get(c.nv));
                }
                // @Mark:=>There is no missing `break`, this is normal logic here
            case INLINESTR:
                if (isNotBlank(c.sv)) {
                    c.sv = c.sv.trim();
                    // 00:00:00
                    if (c.sv.length() == 8 && c.sv.charAt(2) == ':' && c.sv.charAt(5) == ':') {
                        lt = toLocalTime(c.sv);
                    } else {
                        lt = toTimestamp(c.sv).toLocalDateTime().toLocalTime();
                    }
                } else lt = null;
                break;
//            case UNALLOCATED:
//            case BLANK:
//            case EMPTY_TAG:
            default:
                lt = null;
//                break;
//            default: throw new UncheckedTypeException("Can't convert cell value to java.sql.Timestamp");
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
            case SST:
            case INLINESTR:
                type = CellType.STRING;
                break;
            case NUMERIC:
            case CHARACTER:
                type = !styles.fastTestDateFmt(c.xf) ? CellType.INTEGER : CellType.DATE;
                break;
            case LONG:
                type = CellType.LONG;
                break;
            case DOUBLE:
                type = !styles.fastTestDateFmt(c.xf) ? CellType.DOUBLE : CellType.DATE;
                break;
            case BOOL:
                type = CellType.BOOLEAN;
                break;
            case DATETIME:
            case DATE:
            case TIME:
                type = CellType.DATE;
                break;
            case EMPTY_TAG:
            case BLANK:
                type = CellType.BLANK;
                break;
            case UNALLOCATED:
                type = CellType.UNALLOCATED;
                break;

            default: type = CellType.STRING;
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
     * Convert to object, support annotation
     *
     * @param clazz the type of binding
     * @param <T>   the type of return object
     * @return T
     */
    public <T> T to(Class<T> clazz) {
        if (hr == null) {
            hr = asHeader();
            return null;
//            throw new UncheckedTypeException("Lost header row info");
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
//            throw new UncheckedTypeException("Lost header row info");
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
                case SST:
                    if (c.sv == null) {
                        c.setSv(sst.get(c.nv));
                    }
                    // @Mark:=>There is no missing `break`, this is normal logic here
                case INLINESTR:
                    joiner.add(c.sv);
                    break;
                case BOOL:
                    joiner.add(String.valueOf(c.bv));
                    break;
//                case FUNCTION: // convert to inner string
//                    joiner.add("<function>");
//                    break;
                case NUMERIC:
                    if (!styles.fastTestDateFmt(c.xf)) joiner.add(String.valueOf(c.nv));
                    else joiner.add(toLocalDate(c.nv).toString());
                    break;
                case LONG:
                    joiner.add(String.valueOf(c.lv));
                    break;
                case DOUBLE:
                    if (!styles.fastTestDateFmt(c.xf)) joiner.add(String.valueOf(c.dv));
                    else if (c.dv > 1.00001) joiner.add(toTimestamp(c.dv).toString());
                    else joiner.add(toLocalTime(c.dv).toString());
                    break;
                case BLANK:
                case EMPTY_TAG:
                    joiner.add(EMPTY);
                    break;
                default:
                    joiner.add(null);
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
        Map<String, Object> data = new LinkedHashMap<>(hr.lc - hr.fc);
        String[] names = hr.names;
        String key;
        for (int i = hr.fc; i < hr.lc; i++) {
            Cell c = cells[i];
            key = names[i];
            // Ignore null key
            if (key == null) continue;
            switch (c.t) {
                case SST:
                    if (c.sv == null) {
                        c.setSv(sst.get(c.nv));
                    }
                    // @Mark:=>There is no missing `break`, this is normal logic here
                case INLINESTR:
                    data.put(key, c.sv);
                    break;
                case BOOL:
                    data.put(key, c.bv);
                    break;
                case NUMERIC:
                    if (!styles.fastTestDateFmt(c.xf)) data.put(key, c.nv);
                    else data.put(key, toTimestamp(c.nv));
                    break;
                case LONG:
                    data.put(key, c.lv);
                    break;
                case DOUBLE:
                    if (!styles.fastTestDateFmt(c.xf)) data.put(key, c.dv);
                    else if (c.dv > 1.00001) data.put(key, toTimestamp(c.dv));
                    else data.put(key, toLocalTime(c.dv));
                    break;
                case BLANK:
                case EMPTY_TAG:
                    data.put(key, EMPTY);
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
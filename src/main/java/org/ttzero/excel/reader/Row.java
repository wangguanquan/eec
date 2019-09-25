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

package org.ttzero.excel.reader;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.ttzero.excel.entity.style.Styles;
import org.ttzero.excel.util.DateUtil;
import org.ttzero.excel.util.StringUtil;

import java.lang.reflect.InvocationTargetException;
import java.math.BigDecimal;
import java.sql.Timestamp;
import java.util.Date;
import java.util.StringJoiner;

import static org.ttzero.excel.reader.Cell.BLANK;
import static org.ttzero.excel.reader.Cell.BOOL;
import static org.ttzero.excel.reader.Cell.CHARACTER;
import static org.ttzero.excel.reader.Cell.DATE;
import static org.ttzero.excel.reader.Cell.DATETIME;
import static org.ttzero.excel.reader.Cell.DOUBLE;
import static org.ttzero.excel.reader.Cell.FUNCTION;
import static org.ttzero.excel.reader.Cell.INLINESTR;
import static org.ttzero.excel.reader.Cell.LONG;
import static org.ttzero.excel.reader.Cell.NUMERIC;
import static org.ttzero.excel.reader.Cell.SST;
import static org.ttzero.excel.reader.Cell.TIME;
import static org.ttzero.excel.util.DateUtil.toLocalDate;
import static org.ttzero.excel.util.DateUtil.toTimestamp;

/**
 * Create by guanquan.wang at 2019-04-17 11:08
 */
public abstract class Row {
    protected Logger logger = LogManager.getLogger(getClass());
    // Index to row
    int index = -1;
    // Index to first column (zero base)
    int fc = 0;
    // Index to last column (zero base)
    int lc = -1;
    // Share cell
    Cell[] cells;
    /**
     * The Shared String Table
     */
    SharedStrings sst;
    // The header row
    private HeaderRow hr;
    boolean unknownLength;

    /**
     * The global styles
     */
    Styles styles;

    /**
     * The number of row. (zero base)
     *
     * @return int value
     */
    public int getRowNumber() {
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
     * Returns the index of the last column (zero base).
     * The last index of column is increment the max available index
     *
     * @return the last column index
     */
    public int getLastColumnIndex() {
        return lc;
    }

    /**
     * Test unused row (not contains any filled or formatted or value)
     *
     * @return true if unused
     */
    public boolean isEmpty() {
        return lc - fc <= 0;
    }

    private String outOfBoundsMsg(int index) {
        return "Index: " + index + ", Size: " + lc;
    }

    protected void rangeCheck(int index) {
        if (index >= lc || index < 0)
            throw new IndexOutOfBoundsException(outOfBoundsMsg(index));
    }

    /**
     * Returns {@link Cell}
     * @param i the position of cell
     * @return the {@link Cell}
     */
    protected Cell getCell(int i) {
        rangeCheck(i);
        return cells[i];
    }

    /**
     * Search {@link Cell} by column name
     *
     * @param name the column name
     * @return the {@link Cell}
     */
    protected Cell getCell(String name) {
        int i = hr.getIndex(name);
        rangeCheck(i);
        return cells[i];
    }

    /**
     * convert row to header_row
     *
     * @return header Row
     */
    public HeaderRow asHeader() {
        HeaderRow hr = HeaderRow.with(this);
        this.hr = hr;
        return hr;
    }

    Row setHr(HeaderRow hr) {
        this.hr = hr;
        return this;
    }

    /**
     * Get boolean value by column index
     *
     * @param columnIndex the cell index
     * @return boolean
     */
    public boolean getBoolean(int columnIndex) {
        Cell c = getCell(columnIndex);
        return getBoolean(c);
    }

    /**
     * Get boolean value by column name
     *
     * @param columnName the cell name
     * @return boolean
     */
    public boolean getBoolean(String columnName) {
        Cell c = getCell(columnName);
        return getBoolean(c);
    }

    /**
     * Get boolean value
     *
     * @param c the {@link Cell}
     * @return boolean
     */
    protected boolean getBoolean(Cell c) {
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
                v = StringUtil.isNotEmpty(c.sv);
                break;
            case INLINESTR:
                v = StringUtil.isNotEmpty(c.sv);
                break;

            default: v = false;
        }
        return v;
    }

    /**
     * Get byte value by column index
     *
     * @param columnIndex the cell index
     * @return byte
     */
    public byte getByte(int columnIndex) {
        Cell c = getCell(columnIndex);
        return getByte(c);
    }

    /**
     * Get byte value by column name
     *
     * @param columnName the cell name
     * @return byte
     */
    public byte getByte(String columnName) {
        Cell c = getCell(columnName);
        return getByte(c);
    }

    /**
     * Get byte value
     *
     * @param c the {@link Cell}
     * @return byte
     */
    protected byte getByte(Cell c) {
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
            default: throw new UncheckedTypeException("can't convert to byte");
        }
        return b;
    }

    /**
     * Get char value by column index
     *
     * @param columnIndex the cell index
     * @return char
     */
    public char getChar(int columnIndex) {
        Cell c = getCell(columnIndex);
        return getChar(c);
    }

    /**
     * Get char value by column name
     *
     * @param columnName the cell name
     * @return char
     */
    public char getChar(String columnName) {
        Cell c = getCell(columnName);
        return getChar(c);
    }

    /**
     * Get char value
     *
     * @param c the {@link Cell}
     * @return char
     */
    protected char getChar(Cell c) {
        char cc = 0;
        switch (c.t) {
            case SST:
                if (c.sv == null) {
                    c.setSv(sst.get(c.nv));
                }
                String s = c.sv;
                if (StringUtil.isNotEmpty(s)) {
                    cc |= s.charAt(0);
                }
                break;
            case INLINESTR:
                s = c.sv;
                if (StringUtil.isNotEmpty(s)) {
                    cc |= s.charAt(0);
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
            default: throw new UncheckedTypeException("can't convert to char");
        }
        return cc;
    }

    /**
     * Get short value by column index
     *
     * @param columnIndex the cell index
     * @return short
     */
    public short getShort(int columnIndex) {
        Cell c = getCell(columnIndex);
        return getShort(c);
    }

    /**
     * Get short value by column name
     *
     * @param columnName the cell name
     * @return short
     */
    public short getShort(String columnName) {
        Cell c = getCell(columnName);
        return getShort(c);
    }

    /**
     * Get short value
     *
     * @param c the {@link Cell}
     * @return short
     */
    protected short getShort(Cell c) {
        short s = 0;
        switch (c.t) {
            case NUMERIC:
                s |= c.nv;
                break;
            case LONG:
                s |= c.lv;
                break;
            case BOOL:
                s |= c.bv ? 1 : 0;
                break;
            case DOUBLE:
                s |= (int) c.dv;
                break;
            default: throw new UncheckedTypeException("can't convert to short");
        }
        return s;
    }

    /**
     * Get int value by column index
     *
     * @param columnIndex the cell index
     * @return int
     */
    public int getInt(int columnIndex) {
        Cell c = getCell(columnIndex);
        return getInt(c);
    }

    /**
     * Get int value by column name
     *
     * @param columnName the cell name
     * @return int
     */
    public int getInt(String columnName) {
        Cell c = getCell(columnName);
        return getInt(c);
    }

    /**
     * Get int value
     *
     * @param c the {@link Cell}
     * @return int
     */
    protected int getInt(Cell c) {
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
            case BOOL:
                n = c.bv ? 1 : 0;
                break;
            case SST:
                if (c.sv == null) {
                    c.setSv(sst.get(c.nv));
                }
                try {
                    n = Integer.parseInt(c.sv);
                } catch (NumberFormatException e) {
                    throw new UncheckedTypeException("String value " + c.sv + " can't convert to int");
                }
                break;
            case INLINESTR:
                try {
                    n = Integer.parseInt(c.sv);
                } catch (NumberFormatException e) {
                    throw new UncheckedTypeException("String value " + c.sv + " can't convert to int");
                }
                break;

            default: throw new UncheckedTypeException("unknown type");
        }
        return n;
    }

    /**
     * Get long value by column index
     *
     * @param columnIndex the cell index
     * @return long
     */
    public long getLong(int columnIndex) {
        Cell c = getCell(columnIndex);
        return getLong(c);
    }

    /**
     * Get long value by column name
     *
     * @param columnName the cell name
     * @return long
     */
    public long getLong(String columnName) {
        Cell c = getCell(columnName);
        return getLong(c);
    }

    /**
     * Get long value
     *
     * @param c the {@link Cell}
     * @return long
     */
    protected long getLong(Cell c) {
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
                try {
                    l = Long.parseLong(c.sv);
                } catch (NumberFormatException e) {
                    throw new UncheckedTypeException("String value " + c.sv + " can't convert to long");
                }
                break;
            case INLINESTR:
                try {
                    l = Long.parseLong(c.sv);
                } catch (NumberFormatException e) {
                    throw new UncheckedTypeException("String value " + c.sv + " can't convert to long");
                }
                break;
            case BOOL:
                l = c.bv ? 1L : 0L;
                break;
            default: throw new UncheckedTypeException("unknown type");
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
    protected String getString(Cell c) {
        String s;
        switch (c.t) {
            case SST:
                if (c.sv == null) {
                    c.setSv(sst.get(c.nv));
                }
                s = c.sv;
                break;
            case INLINESTR:
                s = c.sv;
                break;
            case BLANK:
                s = StringUtil.EMPTY;
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
     * Get float value by column index
     *
     * @param columnIndex the cell index
     * @return float
     */
    public float getFloat(int columnIndex) {
        return (float) getDouble(columnIndex);
    }

    /**
     * Get float value by column index
     *
     * @param columnName the cell index
     * @return float
     */
    public float getFloat(String columnName) {
        return (float) getDouble(columnName);
    }

    /**
     * Get double value by column index
     *
     * @param columnIndex the cell index
     * @return double
     */
    public double getDouble(int columnIndex) {
        Cell c = getCell(columnIndex);
        return getDouble(c);
    }

    /**
     * Get double value by column name
     *
     * @param columnName the cell name
     * @return double
     */
    public double getDouble(String columnName) {
        Cell c = getCell(columnName);
        return getDouble(c);
    }

    /**
     * Get double value
     *
     * @param c the {@link Cell}
     * @return double
     */
    protected double getDouble(Cell c) {
        double d;
        switch (c.t) {
            case DOUBLE:
                d = c.dv;
                break;
            case NUMERIC:
                d = c.nv;
                break;
            case SST:
                try {
                    d = Double.valueOf(c.sv);
                } catch (NumberFormatException e) {
                    throw new UncheckedTypeException("String value " + c.sv + " can't convert to double");
                }
                break;
            case INLINESTR:
                try {
                    d = Double.valueOf(c.sv);
                } catch (NumberFormatException e) {
                    throw new UncheckedTypeException("String value " + c.sv + " can't convert to double");
                }
                break;

            default: throw new UncheckedTypeException("unknown type");
        }
        return d;
    }

    /**
     * Get decimal value by column index
     *
     * @param columnIndex the cell index
     * @return BigDecimal
     */
    public BigDecimal getDecimal(int columnIndex) {
        Cell c = getCell(columnIndex);
        return getDecimal(c);
    }

    /**
     * Get BigDecimal value by column name
     *
     * @param columnName the cell name
     * @return BigDecimal
     */
    public BigDecimal getDecimal(String columnName) {
        Cell c = getCell(columnName);
        return getDecimal(c);
    }

    /**
     * Get BigDecimal value
     *
     * @param c the {@link Cell}
     * @return BigDecimal
     */
    protected BigDecimal getDecimal(Cell c) {
        BigDecimal bd;
        switch (c.t) {
            case DOUBLE:
                bd = BigDecimal.valueOf(c.dv);
                break;
            case NUMERIC:
                bd = BigDecimal.valueOf(c.nv);
                break;
            default:
                bd = new BigDecimal(c.sv);
        }
        return bd;
    }

    /**
     * Get date value by column index
     *
     * @param columnIndex the cell index
     * @return Date
     */
    public Date getDate(int columnIndex) {
        Cell c = getCell(columnIndex);
        return getDate(c);
    }

    /**
     * Get date value by column name
     *
     * @param columnName the cell name
     * @return Date
     */
    public Date getDate(String columnName) {
        Cell c = getCell(columnName);
        return getDate(c);
    }

    /**
     * Get date value
     *
     * @param c the {@link Cell}
     * @return BigDecimal
     */
    protected Date getDate(Cell c) {
        Date date;
        switch (c.t) {
            case NUMERIC:
                date = DateUtil.toDate(c.nv);
                break;
            case DOUBLE:
                date = DateUtil.toDate(c.dv);
                break;
            case SST:
                if (c.sv == null) {
                    c.setSv(sst.get(c.nv));
                }
                date = DateUtil.toDate(c.sv);
                break;
            case INLINESTR:
                date = DateUtil.toDate(c.sv);
                break;
            default: throw new UncheckedTypeException("");
        }
        return date;
    }

    /**
     * Get timestamp value by column index
     *
     * @param columnIndex the cell index
     * @return java.sql.Timestamp
     */
    public Timestamp getTimestamp(int columnIndex) {
        Cell c = getCell(columnIndex);
        return getTimestamp(c);
    }

    /**
     * Get timestamp value by column name
     *
     * @param columnName the cell name
     * @return java.sql.Timestamp
     */
    public Timestamp getTimestamp(String columnName) {
        Cell c = getCell(columnName);
        return getTimestamp(c);
    }

    /**
     * Get timestamp value
     *
     * @param c the {@link Cell}
     * @return java.sql.Timestamp
     */
    protected Timestamp getTimestamp(Cell c) {
        Timestamp ts;
        switch (c.t) {
            case NUMERIC:
                ts = DateUtil.toTimestamp(c.nv);
                break;
            case DOUBLE:
                ts = DateUtil.toTimestamp(c.dv);
                break;
            case SST:
                if (c.sv == null) {
                    c.setSv(sst.get(c.nv));
                }
                ts = DateUtil.toTimestamp(c.sv);
                break;
            case INLINESTR:
                ts = DateUtil.toTimestamp(c.sv);
                break;
            default: throw new UncheckedTypeException("");
        }
        return ts;
    }

    /**
     * Get time value by column index
     *
     * @param columnIndex the cell index
     * @return java.sql.Time
     */
    public java.sql.Time getTime(int columnIndex) {
        Cell c = getCell(columnIndex);
        if (c.t == DOUBLE) {
            return DateUtil.toTime(c.dv);
        }
        throw new UncheckedTypeException("can't convert to java.sql.Time");
    }

    /**
     * Get time value by column name
     *
     * @param columnName the cell name
     * @return java.sql.Time
     */
    public java.sql.Time getTime(String columnName) {
        Cell c = getCell(columnName);
        if (c.t == DOUBLE) {
            return DateUtil.toTime(c.dv);
        }
        throw new UncheckedTypeException("can't convert to java.sql.Time");
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
    protected CellType getCellType(Cell c) {
        CellType type;
        switch (c.t) {
            case SST:
            case INLINESTR:
                type = CellType.STRING;
                break;
            case NUMERIC:
            case CHARACTER:
                type = !styles.fastTestDateFmt(c.s) ? CellType.INTEGER : CellType.DATE;
                break;
            case LONG:
                type = CellType.LONG;
                break;
            case DOUBLE:
                type = !styles.fastTestDateFmt(c.s) ? CellType.DOUBLE : CellType.DATE;
                break;
            case BOOL:
                type = CellType.BOOLEAN;
                break;
            case DATETIME:
            case DATE:
            case TIME:
                type = CellType.DATE;
                break;
            case BLANK:
                type = CellType.BLANK;
                break;

            default: type = CellType.STRING;
        }
        return type;
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
        } else return (T) this;
    }

    /**
     * Returns the binding type if is bound, otherwise returns Row
     *
     * @param <T> the type of binding
     * @return T
     */
    @SuppressWarnings("unchecked")
    public <T> T geet() {
        if (hr != null && hr.getClazz() != null) {
            T t = hr.getT();
            try {
                hr.put(this, t);
            } catch (IllegalAccessException | InvocationTargetException e) {
                throw new UncheckedTypeException("call set method error.", e);
            }
            return t;
        } else return (T) this;
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
            throw new UncheckedTypeException("Lost header row info");
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
            throw new UncheckedTypeException("Lost header row info");
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
                    joiner.add(c.sv);
                    break;
                case INLINESTR:
                    joiner.add(c.sv);
                    break;
                case BOOL:
                    joiner.add(String.valueOf(c.bv));
                    break;
                case FUNCTION:
                    joiner.add("<function>");
                    break;
                case NUMERIC:
                    if (!styles.fastTestDateFmt(c.s)) joiner.add(String.valueOf(c.nv));
                    else joiner.add(toLocalDate(c.nv).toString());
                    break;
                case LONG:
                    joiner.add(String.valueOf(c.lv));
                    break;
                case DOUBLE:
                    if (!styles.fastTestDateFmt(c.s)) joiner.add(String.valueOf(c.dv));
                    else joiner.add(toTimestamp(c.dv).toString());
                    break;
                case BLANK:
                    joiner.add(StringUtil.EMPTY);
                default:
                    joiner.add(null);
            }
        }
        return joiner.toString();
    }
}

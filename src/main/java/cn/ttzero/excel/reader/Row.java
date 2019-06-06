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

import cn.ttzero.excel.entity.style.Styles;
import cn.ttzero.excel.util.DateUtil;
import cn.ttzero.excel.util.StringUtil;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import java.math.BigDecimal;
import java.sql.Timestamp;
import java.util.Date;
import java.util.StringJoiner;

import static cn.ttzero.excel.reader.Cell.BOOL;
import static cn.ttzero.excel.reader.Cell.NUMERIC;
import static cn.ttzero.excel.reader.Cell.FUNCTION;
import static cn.ttzero.excel.reader.Cell.SST;
import static cn.ttzero.excel.reader.Cell.INLINESTR;
import static cn.ttzero.excel.reader.Cell.LONG;
import static cn.ttzero.excel.reader.Cell.DOUBLE;
import static cn.ttzero.excel.reader.Cell.BLANK;

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
        if (index >= lc)
            throw new IndexOutOfBoundsException(outOfBoundsMsg(index));
    }

    protected Cell getCell(int i) {
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
     * Get double value by column index
     *
     * @param columnIndex the cell index
     * @return double
     */
    public double getDouble(int columnIndex) {
        Cell c = getCell(columnIndex);
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
            } catch (InstantiationException | IllegalAccessException e) {
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
            } catch (IllegalAccessException e) {
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
        } catch (InstantiationException | IllegalAccessException e) {
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
        } catch (IllegalAccessException e) {
            throw new UncheckedTypeException("call set method error.", e);
        }
        return t;
    }

    @Override
    public String toString() {
        if (isEmpty()) return null;
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
                    joiner.add(String.valueOf(c.nv));
                    break;
                case LONG:
                    joiner.add(String.valueOf(c.lv));
                    break;
                case DOUBLE:
                    joiner.add(String.valueOf(c.dv));
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

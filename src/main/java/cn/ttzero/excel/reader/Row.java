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

import cn.ttzero.excel.util.DateUtil;
import cn.ttzero.excel.util.StringUtil;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import java.math.BigDecimal;
import java.sql.Timestamp;
import java.util.Date;
import java.util.StringJoiner;

/**
 * Create by guanquan.wang at 2019-04-17 11:08
 */
public abstract class Row {
    protected Logger logger = LogManager.getLogger(getClass());
    // Index to row
    int index = -1;
    // Index to first column
    int fc = 0;
    // Index to last column
    int lc = -1;
    // Share cell
    Cell[] cells;
    SharedString sst;
    // The header row
    HeaderRow hr;
    boolean unknownLength;

    /**
     * The number of row. (zero base)
     * @return int value
     */
    public int getRowNumber() {
        return index;
    }

    /**
     * Test unused row (not contains any filled or formatted or value)
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
     * @param columnIndex the cell index
     * @return boolean
     */
    public boolean getBoolean(int columnIndex) {
        Cell c = getCell(columnIndex);
        boolean v;
        switch (c.getT()) {
            case 'b':
                v = c.getBv();
                break;
            case 'n':
            case 'd':
                v = c.getNv() != 0 || c.getDv() >= 0.000001 || c.getDv() <= -0.000001;
                break;
            case 's':
                if (c.getSv() == null) {
                    c.setSv(sst.get(c.getNv()));
                }
                v = StringUtil.isNotEmpty(c.getSv());
                break;
            case 'r':
                v = StringUtil.isNotEmpty(c.getSv());
                break;

            default: v = false;
        }
        return v;
    }

    /**
     * Get byte value by column index
     * @param columnIndex the cell index
     * @return byte
     */
    public byte getByte(int columnIndex) {
        Cell c = getCell(columnIndex);
        byte b = 0;
        switch (c.getT()) {
            case 'n':
                b |= c.getNv();
                break;
            case 'l':
                b |= c.getLv();
                break;
            case 'b':
                b |= c.getBv() ? 1 : 0;
                break;
            case 'd':
                b |= (int) c.getDv();
                break;
            default: throw new UncheckedTypeException("can't convert to byte");
        }
        return b;
    }

    /**
     * Get char value by column index
     * @param columnIndex the cell index
     * @return char
     */
    public char getChar(int columnIndex) {
        Cell c = getCell(columnIndex);
        char cc = 0;
        switch (c.getT()) {
            case 's':
                if (c.getSv() == null) {
                    c.setSv(sst.get(c.getNv()));
                }
                String s = c.getSv();
                if (StringUtil.isNotEmpty(s)) {
                    cc |= s.charAt(0);
                }
                break;
            case 'r':
                s = c.getSv();
                if (StringUtil.isNotEmpty(s)) {
                    cc |= s.charAt(0);
                }
                break;
            case 'n':
                cc |= c.getNv();
                break;
            case 'l':
                cc |= c.getLv();
                break;
            case 'b':
                cc |= c.getBv() ? 1 : 0;
                break;
            case 'd':
                cc |= (int) c.getDv();
                break;
            default: throw new UncheckedTypeException("can't convert to char");
        }
        return cc;
    }

    /**
     * Get short value by column index
     * @param columnIndex the cell index
     * @return short
     */
    public short getShort(int columnIndex) {
        Cell c = getCell(columnIndex);
        short s = 0;
        switch (c.getT()) {
            case 'n':
                s |= c.getNv();
                break;
            case 'l':
                s |= c.getLv();
                break;
            case 'b':
                s |= c.getBv() ? 1 : 0;
                break;
            case 'd':
                s |= (int) c.getDv();
                break;
            default: throw new UncheckedTypeException("can't convert to short");
        }
        return s;
    }

    /**
     * Get int value by column index
     * @param columnIndex the cell index
     * @return int
     */
    public int getInt(int columnIndex) {
        Cell c = getCell(columnIndex);
        int n;
        switch (c.getT()) {
            case 'n':
                n = c.getNv();
                break;
            case 'l':
                n = (int) c.getLv();
                break;
            case 'd':
                n = (int) c.getDv();
                break;
            case 'b':
                n = c.getBv() ? 1 : 0;
                break;
            case 's':
                if (c.getSv() == null) {
                    c.setSv(sst.get(c.getNv()));
                }
                try {
                    n = Integer.parseInt(c.getSv());
                } catch (NumberFormatException e) {
                    throw new UncheckedTypeException("String value " + c.getSv() + " can't convert to int");
                }
                break;
            case 'r':
                try {
                    n = Integer.parseInt(c.getSv());
                } catch (NumberFormatException e) {
                    throw new UncheckedTypeException("String value " + c.getSv() + " can't convert to int");
                }
                break;

            default: throw new UncheckedTypeException("unknown type");
        }
        return n;
    }

    /**
     * Get long value by column index
     * @param columnIndex the cell index
     * @return long
     */
    public long getLong(int columnIndex) {
        Cell c = getCell(columnIndex);
        long l;
        switch (c.getT()) {
            case 'l':
                l = c.getLv();
                break;
            case 'n':
                l = c.getNv();
                break;
            case 'd':
                l = (long) c.getDv();
                break;
            case 's':
                if (c.getSv() == null) {
                    c.setSv(sst.get(c.getNv()));
                }
                try {
                    l = Long.parseLong(c.getSv());
                } catch (NumberFormatException e) {
                    throw new UncheckedTypeException("String value " + c.getSv() + " can't convert to long");
                }
                break;
            case 'r':
                try {
                    l = Long.parseLong(c.getSv());
                } catch (NumberFormatException e) {
                    throw new UncheckedTypeException("String value " + c.getSv() + " can't convert to long");
                }
                break;
            case 'b':
                l = c.getBv() ? 1L : 0L;
                break;
            default: throw new UncheckedTypeException("unknown type");
        }
        return l;
    }

    /**
     * Get string value by column index
     * @param columnIndex the cell index
     * @return string
     */
    public String getString(int columnIndex) {
        Cell c = getCell(columnIndex);
        String s;
        switch (c.getT()) {
            case 's':
                if (c.getSv() == null) {
                    c.setSv(sst.get(c.getNv()));
                }
                s = c.getSv();
                break;
            case 'r':
                s = c.getSv();
                break;
            case 'l':
                s = String.valueOf(c.getLv());
                break;
            case 'n':
                s = String.valueOf(c.getNv());
                break;
            case 'd':
                s = String.valueOf(c.getDv());
                break;
            case 'b':
                s = c.getBv() ? "true" : "false";
                break;
            default: s = c.getSv();
        }
        return s;
    }

    /**
     * Get float value by column index
     * @param columnIndex the cell index
     * @return float
     */
    public float getFloat(int columnIndex) {
        return (float) getDouble(columnIndex);
    }

    /**
     * Get double value by column index
     * @param columnIndex the cell index
     * @return double
     */
    public double getDouble(int columnIndex) {
        Cell c = getCell(columnIndex);
        double d;
        switch (c.getT()) {
            case 'd':
                d = c.getDv();
                break;
            case 'n':
                d = c.getNv();
                break;
            case 's':
                try {
                    d = Double.valueOf(c.getSv());
                } catch (NumberFormatException e) {
                    throw new UncheckedTypeException("String value " + c.getSv() + " can't convert to double");
                }
                break;
            case 'r':
                try {
                    d = Double.valueOf(c.getSv());
                } catch (NumberFormatException e) {
                    throw new UncheckedTypeException("String value " + c.getSv() + " can't convert to double");
                }
                break;

            default: throw new UncheckedTypeException("unknown type");
        }
        return d;
    }

    /**
     * Get decimal value by column index
     * @param columnIndex the cell index
     * @return BigDecimal
     */
    public BigDecimal getDecimal(int columnIndex) {
        Cell c = getCell(columnIndex);
        BigDecimal bd;
        switch (c.getT()) {
            case 'd':
                bd = BigDecimal.valueOf(c.getDv());
                break;
            case 'n':
                bd = BigDecimal.valueOf(c.getNv());
                break;
            default:
                bd = new BigDecimal(c.getSv());
        }
        return bd;
    }

    /**
     * Get date value by column index
     * @param columnIndex the cell index
     * @return Date
     */
    public Date getDate(int columnIndex) {
        Cell c = getCell(columnIndex);
        Date date;
        switch (c.getT()) {
            case 'n':
                date = DateUtil.toDate(c.getNv());
                break;
            case 'd':
                date = DateUtil.toDate(c.getDv());
                break;
            case 's':
                if (c.getSv() == null) {
                    c.setSv(sst.get(c.getNv()));
                }
                date = DateUtil.toDate(c.getSv());
                break;
            case 'r':
                date = DateUtil.toDate(c.getSv());
                break;
            default: throw new UncheckedTypeException("");
        }
        return date;
    }

    /**
     * Get timestamp value by column index
     * @param columnIndex the cell index
     * @return java.sql.Timestamp
     */
    public Timestamp getTimestamp(int columnIndex) {
        Cell c = getCell(columnIndex);
        Timestamp ts;
        switch (c.getT()) {
            case 'n':
                ts = DateUtil.toTimestamp(c.getNv());
                break;
            case 'd':
                ts = DateUtil.toTimestamp(c.getDv());
                break;
            case 's':
                if (c.getSv() == null) {
                    c.setSv(sst.get(c.getNv()));
                }
                ts = DateUtil.toTimestamp(c.getSv());
                break;
            case 'r':
                ts = DateUtil.toTimestamp(c.getSv());
                break;
            default: throw new UncheckedTypeException("");
        }
        return ts;
    }

    /**
     * Get time value by column index
     * @param columnIndex the cell index
     * @return java.sql.Time
     */
    public java.sql.Time getTime(int columnIndex) {
        Cell c = getCell(columnIndex);
        if (c.getT() == 'd') {
            return DateUtil.toTime(c.getDv());
        }
        throw new UncheckedTypeException("can't convert to java.sql.Time");
    }

    /**
     * Returns the binding type if is bound, otherwise returns Row
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
     * @param <T> the type of binding
     * @return T
     */
    @SuppressWarnings("unchecked")
    public <T> T geet() {
        if (hr != null && hr.getClazz() != null) {
            T t = hr.getT();
            try {
                hr.put(this, t);
            } catch (IllegalAccessException  e) {
                throw new UncheckedTypeException("call set method error.", e);
            }
            return t;
        } else return (T) this;
    }
    /////////////////////////////To object//////////////////////////////////

    /**
     * Convert to object, support annotation
     * @param clazz the type of binding
     * @param <T> the type of return object
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
     * @param clazz the type of binding
     * @param <T> the type of return object
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
        } catch (IllegalAccessException  e) {
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
        for (int i = fc - 1; i < lc; i++) {
            Cell c = cells[i];
            switch (c.getT()) {
                case 's':
                    if (c.getSv() == null) {
                        c.setSv(sst.get(c.getNv()));
                    }
                    joiner.add(c.getSv());
                    break;
                case 'r':
                    joiner.add(c.getSv());
                    break;
                case 'b':
                    joiner.add(String.valueOf(c.getBv()));
                    break;
                case 'f':
                    joiner.add("<function>");
                    break;
                case 'n':
                    joiner.add(String.valueOf(c.getNv()));
                    break;
                case 'l':
                    joiner.add(String.valueOf(c.getLv()));
                    break;
                case 'd':
                    joiner.add(String.valueOf(c.getDv()));
                    break;
                default:
                    joiner.add(null);
            }
        }
        return joiner.toString();
    }
}

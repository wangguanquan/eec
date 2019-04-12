/*
 * Copyright (c) 2009, guanquan.wang@yandex.com All Rights Reserved.
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

import cn.ttzero.excel.annotation.DisplayName;
import cn.ttzero.excel.annotation.NotImport;
import cn.ttzero.excel.util.DateUtil;
import cn.ttzero.excel.util.StringUtil;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.sql.Timestamp;
import java.util.Arrays;
import java.util.Date;
import java.util.StringJoiner;

/**
 * 行数据，同一个Sheet页内的Row对象内存共享。
 * 行数据开始列和结束列读取的是span值，你可以使用<code>row.isEmpty()</code>方法测试行数据是否为空节点
 * 空节点定义为: 没有任何值和样式以及格式化的行. 像这样<code><row r="x"/></code>
 * 你可以像ResultSet一样通过单元格下标获取数据eq:<code>row.getInt(1) // 获取当前行第2列的数据</code>下标从0开始。
 * 也可以使用to&amp;too方法将行数据转为对象，前者会实例化每个对象，后者内存共享只会有一个实例,在流式操作中这是一个好主意。
 *
 * Create by guanquan.wang on 2018-09-22
 */
public class Row {
    protected Logger logger = LogManager.getLogger(getClass());
    private int rowNumber = -1, p2 = -1, p1 = 0;
    private Cell[] cells;
    private SharedString sst;
    private HeaderRow hr;
    private int startRow;
    private boolean unknownLength;

    public int getRowNumber() {
        if (rowNumber == -1)
            searchRowNumber();
        return rowNumber;
    }

    private Row(){}

    Row(SharedString sst, int startRow) {
        this.sst = sst;
        this.startRow = startRow;
    }

    /////////////////////////unsafe////////////////////////
    private char[] cb;
    private int from, to;
    private int cursor;
    ///////////////////////////////////////////////////////
    Row with(char[] cb, int from, int size) {
//        logger.info(new String(cb, from, size));
        this.cb = cb;
        this.from = from;
        this.to = from + size;
        this.cursor = from;
        this.rowNumber = this.p2 = -1;
        parseCells();
        return this;
    }

    /* empty row*/
    Row empty(char[] cb, int from, int size) {
//        logger.info(new String(cb, from, size));
        this.cb = cb;
        this.from = from;
        this.to = from + size;
        this.cursor = from;
        this.rowNumber = -1;
        this.p1 = this.p2 = -1;
        return this;
    }

    private int searchRowNumber() {
        int _f = from + 4, a; // skip '<row '
        for (; cb[_f] != '>' && _f < to; _f++) {
            if (cb[_f] == ' ' && cb[_f + 1] == 'r' && cb[_f + 2] == '=') {
                a = _f += 4;
                for (; cb[_f] != '"' && _f < to; _f++);
                if (_f > a) {
                    rowNumber = toInt(a, _f);
                }
                break;
            }
        }
        return _f;
    }

    private int searchSpan() {
        int i = from;
        for (; cb[i] != '>'; i++) {
            if (cb[i] == ' ' && cb[i + 1] == 's' && cb[i + 2] == 'p'
                    && cb[i + 3] == 'a' && cb[i + 4] == 'n' && cb[i + 5] == 's'
                    && cb[i + 6] == '=') {
                i += 8;
                int b, j = i;
                for (; cb[i] != '"' && cb[i] != '>'; i++);
                for (b = i - 1; cb[b] != ':'; b--);
                if (++b < i) {
                    p2 = toInt(b, i);
                }
                if (j < --b) {
                    p1 = toInt(j, b);
                }
            }
        }
        if (p1 <= 0) p1 = this.startRow;
        if (cells == null || p2 > cells.length) {
            cells = new Cell[p2 > 0 ? p2 : 100]; // default array length 100
        }
        // clear and share
        for (int n = 0, len = p2 > 0 ? p2 : cells.length; n < len; n++) {
            if (cells[n] != null) cells[n].clear();
            else cells[n] = new Cell();
        }
        return i;
    }

    /**
     * 解析每行数据
     */
    private void parseCells() {
        int index = 0;
        cursor = searchSpan();
        for (; cb[cursor++] != '>'; );
        unknownLength = p2 < 0;
        if (unknownLength) {
            while (nextCell() != null) index++;
        } else {
            while (index < p2 && nextCell() != null);
        }
    }

    /**
     * 迭代每列数据
     * @return Cell
     */
    protected Cell nextCell() {
        for (; cursor < to && (cb[cursor] != '<' || cb[cursor + 1] != 'c' || cb[cursor + 2] != ' '); cursor++);
        // end of row
        if (cursor >= to) return null;
        cursor += 2;
        // find end of cell
        int e = cursor;
        for (; e < to && (cb[e] != '<' || cb[e + 1] != 'c' || cb[e + 2] != ' '); e++);

        Cell cell = null;
        // find type
        // n=numeric (default), s=string, b=boolean, str=function string
        char t = 'n'; // default
        for (; cb[cursor] != '>'; cursor++) {
            // cell index
            if (cb[cursor] == ' ' && cb[cursor + 1] == 'r' && cb[cursor + 2] == '=') {
                int a = cursor += 4;
                for (; cb[cursor] != '"'; cursor++);
                cell = cells[unknownLength ? (p2 = toCellIndex(a, cursor)) - 1 : toCellIndex(a, cursor) - 1];
            }
            // cell type
            if (cb[cursor] == ' ' && cb[cursor + 1] == 't' && cb[cursor + 2] == '=') {
                int a = cursor += 4, n;
                for (; cb[cursor] != '"'; cursor++);
                if ((n = cursor - a) == 1) {
                    t = cb[a]; // s, n, b
                } else if (n == 3 && cb[a] == 's' && cb[a + 1] == 't' && cb[a + 2] == 'r') {
                    t = 'f'; // function string
                } else if (n == 9 && cb[a] == 'i' && cb[a + 1] == 'n' && cb[a + 2] == 'l' && cb[a + 6] == 'S' && cb[a + 8] == 'r') {
                    t = 'r'; // inlineStr
                }
                // -> other unknown case
            }
        }

        if (cell == null) return null;

        cell.setT(t);

        // get value
        int a;
        switch (t) {
            case 'r': // inner string
                a = getT(e);
                if (a == cursor) { // null value
                    cell.setSv(null);
                } else {
                    cell.setSv(sst.unescape(cb, a, cursor));
                }
//                cell.setT('s'); // reset type to string
                break;
            case 's': // shared string lazy get
                a = getV(e);
                cell.setNv(toInt(a, cursor));
//                cell.setSv(sst.get(toInt(a, cursor)));
                break;
            case 'b': // boolean value
                a = getV(e);
                if (cursor - a == 1) {
                    cell.setBv(toInt(a, cursor) == 1);
                }
                break;
            case 'f': // function string
                break;
            default:
                a = getV(e);
                if (a < cursor) {
                    if (isNumber(a, cursor)) {
                        long l = toLong(a, cursor);
                        if (l <= Integer.MAX_VALUE && l >= Integer.MIN_VALUE) {
                            cell.setNv((int) l);
                            cell.setT('n');
                        } else {
                            cell.setLv(l);
                            cell.setT('l');
                        }
                    } else if (isDouble(a, cursor)) {
                        cell.setT('d');
                        cell.setDv(toDouble(a, cursor));
                    } else {
                        cell.setT('s');
                        cell.setSv(toString(a, cursor));
                    }
                }
        }

        // end of cell
        cursor = e;

        return cell;
    }

    private int toInt(int a, int b) {
        boolean _n;
        if (_n = cb[a] == '-') a++;
        int n = cb[a++] - '0';
        for (; b > a; ) {
            n = n * 10 + cb[a++] - '0';
        }
        return _n ? -n : n;
    }

    private long toLong(int a, int b) {
        boolean _n;
        if (_n = cb[a] == '-') a++;
        long n = cb[a++] - '0';
        for (; b > a; ) {
            n = n * 10 + cb[a++] - '0';
        }
        return _n ? -n : n;
    }

    private String toString(int a, int b) {
        return new String(cb, a, b - a);
    }

    private double toDouble(int a, int b) {
        return Double.valueOf(toString(a, b));
    }

    private boolean isNumber(int a, int b) {
        if (a == b) return false;
        if (cb[a] == '-') a++;
        for ( ; a < b; ) {
            char c = cb[a++];
            if (c < '0' || c > '9') break;
        }
        return a == b;
    }

    /**
     * FIXME check double
     * @param a
     * @param b
     * @return
     */
    private boolean isDouble(int a, int b) {
        if (a == b) return false;
        if (cb[a] == '-') a++;
        for (char i = 0 ; a < b; ) {
            char c = cb[a++];
            if (i > 1) return false;
            if (c >= '0' && c <= '9') continue;
            if (c == '.') i++;
        }
        return true;
    }

    /**
     * inner string
     * <is><t>cell value</t></is>
     * @param e
     * @return
     */
    private int getT(int e) {
        for (; cursor < e && (cb[cursor] != '<' || cb[cursor + 1] != 't' || cb[cursor + 2] != '>'); cursor++);
        if (cursor == e) return cursor;
        int a = cursor += 3;
        for (; cursor < e && (cb[cursor] != '<' || cb[cursor + 1] != '/' || cb[cursor + 2] != 't' || cb[cursor + 3] != '>'); cursor++);
        return a;
    }

    /**
     * shared string
     * <v>1</v>
     * @param e
     * @return
     */
    private int getV(int e) {
        for (; cursor < e && (cb[cursor] != '<' || cb[cursor + 1] != 'v' || cb[cursor + 2] != '>'); cursor++);
        if (cursor == e) return cursor;
        int a = cursor += 3;
        for (; cursor < e && (cb[cursor] != '<' || cb[cursor + 1] != '/' || cb[cursor + 2] != 'v' || cb[cursor + 3] != '>'); cursor++);
        return a;
    }

    /**
     * function string
     * @param e
     * @return
     */
    private int getF(int e) {
        // undo
        // return end index of row
        return e;
    }

    /**
     * 26进制转10进制
     * @param a
     * @param b
     * @return
     */
    private int toCellIndex(int a, int b) {
        int n = 0;
        for ( ; a <= b; a++) {
            if (cb[a] <= 'Z' && cb[a] >= 'A') {
                n = n * 26 + cb[a] - '@';
            } else if (cb[a] <= 'z' && cb[a] >= 'a') {
                n = n * 26 + cb[a] - '、';
            } else break;
        }
        return n;
    }

    /**
     * 是否为空行，行不带任何字符和格式内容
     * @return boolean
     */
    public boolean isEmpty() {
        return p1 == -1 && p1 == p2;
    }

    @Override public String toString() {
        if (isEmpty()) return null;
        StringJoiner joiner = new StringJoiner(" | ");
        // show row number
//        joiner.add(String.valueOf(getRowNumber()));
        for (int i = p1 - 1; i < p2; i++) {
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

    /**
     * convert row to header_row
     * @return header Row
     */
    HeaderRow asHeader() {
        HeaderRow hr = HeaderRow.with(this);
        this.hr = hr;
        return hr;
    }

    Row setHr(HeaderRow hr) {
        this.hr = hr;
        return this;
    }

    //////////////////////////////////////Read Value///////////////////////////////////
    private String outOfBoundsMsg(int index) {
        return "Index: " + index + ", Size: " + p2;
    }
    protected void rangeCheck(int index) {
        if (index >= p2)
            throw new IndexOutOfBoundsException(outOfBoundsMsg(index));
    }

    protected Cell getCell(int i) {
        rangeCheck(i);
        return cells[i];
    }

    /**
     * get by index
     * @param columnIndex cell index
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
     * get by index
     * @param columnIndex cell index
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
     * get by index
     * @param columnIndex cell index
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
     * get by index
     * @param columnIndex cell index
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
     * get by index
     * @param columnIndex cell index
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
     * get by index
     * @param columnIndex cell index
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
     * get by index
     * @param columnIndex cell index
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
     * get by index
     * @param columnIndex cell index
     * @return float
     */
    public float getFloat(int columnIndex) {
        return (float) getDouble(columnIndex);
    }
    /**
     * get by index
     * @param columnIndex cell index
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
     * get by index
     * @param columnIndex cell index
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
     * get by index
     * @param columnIndex cell index
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
     * get by index
     * @param columnIndex cell index
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
     * get by index
     * @param columnIndex cell index
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
     * override this method
     * @param columnIndex
     * @param <T>
     * @return
     */
    public <T> T get(int columnIndex) {
        throw new UnsupportedOperationException();
    }

    /**
     * 如果绑定了类型则返回绑定类型否则返回Row
     * @param <T> 指定类型
     * @return 实例
     */
    public <T> T get() {
        if (hr != null && hr.clazz != null) {
            T t;
            try {
                t = (T) hr.clazz.newInstance();
                put(t);
            } catch (InstantiationException | IllegalAccessException e) {
                throw new UncheckedTypeException(hr.clazz + " new instance error.", e);
            }
            return t;
        } else return (T) this;
    }

    /**
     * 如果绑定了类型则返回绑定类型否则返回Row
     * @param <T> 指定类型
     * @return 共享实例
     */
    public <T> T geet() {
        if (hr != null && hr.clazz != null) {
            T t = hr.getT();
            try {
                put(t);
            } catch (IllegalAccessException  e) {
                throw new UncheckedTypeException("call set method error.", e);
            }
            return t;
        } else return (T) this;
    }
    /////////////////////////////To object//////////////////////////////////

    /**
     * convert to object, support annotation
     * @param clazz
     * @param <T>
     * @return
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
            put(t);
        } catch (InstantiationException | IllegalAccessException e) {
            throw new UncheckedTypeException(clazz + " new instance error.", e);
        }
        return t;
    }

    /**
     * memory shared object
     * @param clazz convert to class
     * @param <T> class
     * @return
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
            put(t);
        } catch (IllegalAccessException  e) {
            throw new UncheckedTypeException("call set method error.", e);
        }
        return t;
    }

    private void put(Object t) throws IllegalAccessException {
        int[] columns = hr.getColumns();
        Field[] fields = hr.getFields();
        Class<?>[] fieldClazz = hr.getFieldClazz();
        for (int i = 0; i < columns.length; i++) {
            int c = columns[i];
            if (fieldClazz[i] == String.class) {
                fields[i].set(t, getString(c));
            }
            else if (fieldClazz[i] == int.class || fieldClazz[i] == Integer.class) {
                fields[i].setInt(t, getInt(c));
            }
            else if (fieldClazz[i] == long.class || fieldClazz[i] == Long.class) {
                fields[i].setLong(t, getLong(c));
            }
            else if (fieldClazz[i] == java.util.Date.class || fieldClazz[i] == java.sql.Date.class) {
                fields[i].set(t, getDate(c));
            }
            else if (fieldClazz[i] == java.sql.Timestamp.class) {
                fields[i].set(t, getTimestamp(c));
            }
            else if (fieldClazz[i] == double.class || fieldClazz[i] == Double.class) {
                fields[i].setDouble(t, getDouble(c));
            }
            else if (fieldClazz[i] == float.class || fieldClazz[i] == Float.class) {
                fields[i].setFloat(t, getFloat(c));
            }
            else if (fieldClazz[i] == boolean.class || fieldClazz[i] == Boolean.class) {
                fields[i].setBoolean(t, getBoolean(c));
            }
            else if (fieldClazz[i] == char.class || fieldClazz[i] == Character.class) {
                fields[i].setChar(t, getChar(c));
            }
            else if (fieldClazz[i] == byte.class || fieldClazz[i] == Byte.class) {
                fields[i].setByte(t, getByte(c));
            }
            else if (fieldClazz[i] == short.class || fieldClazz[i] == Short.class) {
                fields[i].setShort(t, getShort(c));
            }
        }
    }

    ////////////////////////private inner class///////////////////////////////
    static final class HeaderRow extends Row {
        String[] names;
        Class<?> clazz;
        Field[] fields;
        int[] columns;
        Class<?>[] fieldClazz;
        Object t;

        static final HeaderRow with(Row row) {
            HeaderRow hr = new HeaderRow();
            hr.names = new String[row.p2];
            Cell c;
            for (int i = row.p1 - 1; i < row.p2; i++) {
                c = row.cells[i];
                // header type is string
                if (c.getT() == 's') {
                    hr.names[i] = row.sst.get(c.getNv());
                } else {
                    hr.names[i] = c.getSv();
                }
            }
            return hr;
        }

        final boolean is(Class<?> clazz) {
            return this.clazz != null && this.clazz == clazz;
        }

        /**
         * mapping
         * @param clazz
         * @return
         */
        final HeaderRow setClass(Class<?> clazz) {
            this.clazz = clazz;
            Field[] fields = clazz.getDeclaredFields();
            int[] index = new int[fields.length];
            int count = 0;
            for (int i = 0, n; i < fields.length; i++) {
                Field f = fields[i];
                // skip not import fields
                NotImport nit = f.getAnnotation(NotImport.class);
                if (nit != null) {
                    fields[i] = null;
                    continue;
                }
                // field has display name
                DisplayName ano = f.getAnnotation(DisplayName.class);
                if (ano != null && StringUtil.isNotEmpty(ano.value())) {
                    n = StringUtil.indexOf(names, ano.value());
                    if (n == -1) {
                        logger.warn(clazz + " field [" + ano.value() + "] can't find in header" + Arrays.toString(names));
                        fields[i] = null;
                        continue;
                    }
                }
                // no annotation or annotation value is null
                else {
                    String name = f.getName();
                    n = StringUtil.indexOf(names, name);
                    if (n == -1 && (n = StringUtil.indexOf(names, StringUtil.toPascalCase(name))) == -1) {
                        fields[i] = null;
                        continue;
                    }
                }

                index[i] = n;
                count++;
            }

            this.fields = new Field[count];
            this.columns = new int[count];
            this.fieldClazz = new Class<?>[count];

            for (int i = fields.length - 1; i >= 0; i--) {
                if (fields[i] != null) {
                    count--;
                    this.fields[count] = fields[i];
                    this.fields[count].setAccessible(true);
                    this.columns[count] = index[i];
                    this.fieldClazz[count] = fields[i].getType();
                }
            }
            return this;
        }

        /**
         * mapping and instance
         * @param clazz
         * @return
         * @throws IllegalAccessException
         * @throws InstantiationException
         */
        final HeaderRow setClassOnce(Class<?> clazz) throws IllegalAccessException, InstantiationException {
            setClass(clazz);
            this.t = clazz.newInstance();
            return this;
        }

        final Field[] getFields() {
            return fields;
        }

        final int[] getColumns() {
            return columns;
        }

        final Class<?>[] getFieldClazz() {
            return fieldClazz;
        }

        final <T> T getT() {
            return (T) t;
        }

        @Override public String get(int columnIndex) {
            rangeCheck(columnIndex);
            return names[columnIndex];
        }

        @Override public String toString() {
            StringJoiner joiner = new StringJoiner(" | ");
            int i = 0;
            for (; names[i++] == null; );
            for (; i < names.length; i++) {
                joiner.add(names[i]);
            }
            return joiner.toString();
        }
    }
}

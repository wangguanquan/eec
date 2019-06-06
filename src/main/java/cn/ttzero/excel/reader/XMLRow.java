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

import static cn.ttzero.excel.reader.Cell.BOOL;
import static cn.ttzero.excel.reader.Cell.NUMERIC;
import static cn.ttzero.excel.reader.Cell.FUNCTION;
import static cn.ttzero.excel.reader.Cell.SST;
import static cn.ttzero.excel.reader.Cell.INLINESTR;
import static cn.ttzero.excel.reader.Cell.BLANK;
import static cn.ttzero.excel.reader.SharedStrings.toInt;
import static cn.ttzero.excel.reader.SharedStrings.unescape;

/**
 * 行数据，同一个Sheet页内的Row对象内存共享。
 * 行数据开始列和结束列读取的是span值，你可以使用<code>row.isEmpty()</code>方法测试行数据是否为空节点
 * 空节点定义为: 没有任何值和样式以及格式化的行. 像这样<code><row r="x"/></code>
 * 你可以像ResultSet一样通过单元格下标获取数据eq:<code>row.getInt(1) // 获取当前行第2列的数据</code>下标从0开始。
 * 也可以使用to&amp;too方法将行数据转为对象，前者会实例化每个对象，后者内存共享只会有一个实例,在流式操作中这是一个好主意。
 * <p>
 * Create by guanquan.wang on 2018-09-22
 */
class XMLRow extends Row {
    private int startRow;
    private StringBuilder buf;

    /**
     * The number of row. (zero base)
     *
     * @return int value
     */
    @Override
    public int getRowNumber() {
        if (index == -1)
            searchRowNumber();
        // The first row index is one
        return index - 1;
    }

    private XMLRow() { }

    XMLRow(SharedStrings sst, Styles styles, int startRow) {
        this.sst = sst;
        this.styles = styles;
        this.startRow = startRow;
        buf = new StringBuilder();
    }

    /////////////////////////unsafe////////////////////////
    private char[] cb;
    private int from, to;
    private int cursor;

    ///////////////////////////////////////////////////////
    XMLRow with(char[] cb, int from, int size) {
//        logger.info(new String(cb, from, size));
        this.cb = cb;
        this.from = from;
        this.to = from + size;
        this.cursor = from;
        this.index = this.lc = -1;
        parseCells();
        return this;
    }

    /* empty row*/
    XMLRow empty(char[] cb, int from, int size) {
//        logger.info(new String(cb, from, size));
        this.cb = cb;
        this.from = from;
        this.to = from + size;
        this.cursor = from;
        this.index = -1;
        this.fc = this.lc = -1;
        return this;
    }

    private int searchRowNumber() {
        int _f = from + 4, a; // skip '<row '
        for (; cb[_f] != '>' && _f < to; _f++) {
            if (cb[_f] == ' ' && cb[_f + 1] == 'r' && cb[_f + 2] == '=') {
                a = _f += 4;
                for (; cb[_f] != '"' && _f < to; _f++) ;
                if (_f > a) {
                    index = toInt(cb, a, _f);
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
                for (; cb[i] != '"' && cb[i] != '>'; i++) ;
                for (b = i - 1; cb[b] != ':'; b--) ;
                if (++b < i) {
                    lc = toInt(cb, b, i);
                }
                if (j < --b) {
                    fc = toInt(cb, j, b);
                }
            }
        }
        if (fc <= 0) fc = this.startRow;
        fc = fc - 1; // zero base
        if (cells == null || lc > cells.length) {
            cells = new Cell[lc > 0 ? lc : 100]; // default array length 100
        }
        // clear and share
        for (int n = 0, len = lc > 0 ? lc : cells.length; n < len; n++) {
            if (cells[n] != null) cells[n].clear();
            else cells[n] = new Cell();
        }
        return i;
    }

    /**
     * Loop parse cell
     */
    private void parseCells() {
        int index = 0;
        cursor = searchSpan();
        for (; cb[cursor++] != '>'; ) ;
        unknownLength = lc < 0;
        if (unknownLength) {
            while (nextCell() != null) index++;
        } else {
            while (index < lc && nextCell() != null) ;
        }
    }

    /**
     * Loop parse cell
     *
     * @return the {@link Cell}
     */
    protected Cell nextCell() {
        for (; cursor < to && (cb[cursor] != '<' || cb[cursor + 1] != 'c'
            || cb[cursor + 2] != ' '); cursor++) ;
        // end of row
        if (cursor >= to) return null;
        cursor += 2;
        // find end of cell
        int e = cursor;
        for (; e < to && (cb[e] != '<' || cb[e + 1] != 'c' || cb[e + 2] != ' '); e++) ;

        Cell cell = null;
        // find type
        // n=numeric (default), s=string, b=boolean, str=function string
        char t = NUMERIC; // default
        // The style index
        short s = 0;
        for (; cb[cursor] != '>'; cursor++) {
            // Cell index
            if (cb[cursor] == ' ' && cb[cursor + 1] == 'r' && cb[cursor + 2] == '=') {
                int a = cursor += 4;
                for (; cb[cursor] != '"'; cursor++) ;
                cell = cells[unknownLength ? (lc = toCellIndex(a, cursor)) - 1 : toCellIndex(a, cursor) - 1];
            }
            // Cell type
            if (cb[cursor] == ' ' && cb[cursor + 1] == 't' && cb[cursor + 2] == '=') {
                int a = cursor += 4, n;
                for (; cb[cursor] != '"'; cursor++) ;
                if ((n = cursor - a) == 1) {
                    t = cb[a]; // s, n, b
                } else if (n == 3 && cb[a] == 's' && cb[a + 1] == 't' && cb[a + 2] == 'r') {
                    t = FUNCTION; // function string
                } else if (n == 9 && cb[a] == 'i' && cb[a + 1] == 'n'
                    && cb[a + 2] == 'l' && cb[a + 6] == 'S' && cb[a + 8] == 'r') {
                    t = INLINESTR; // inlineStr
                }
                // -> other unknown case
            }
            // Cell style
            if (cb[cursor] == ' ' && cb[cursor + 1] == 's' && cb[cursor + 2] == '=') {
                int a = cursor += 4;
                for (; cb[cursor] != '"'; cursor++) ;
                s = (short) (toInt(cb, a, cursor) & 0xFFFF);
            }
        }

        if (cell == null) return null;

        // The style index
        cell.s = s;

        // get value
        int a;
        switch (t) {
            case INLINESTR: // inner string
                a = getT(e);
                if (a == cursor) { // null value
                    cell.setT(BLANK); // Reset type to BLANK if null value
                } else {
                    cell.setSv(unescape(buf, cb, a, cursor));
                }
                break;
            case SST: // shared string lazy get
                a = getV(e);
                cell.setNv(toInt(cb, a, cursor));
                cell.setT(SST);
                break;
            case BOOL: // boolean value
                a = getV(e);
                if (cursor - a == 1) {
                    cell.setBv(toInt(cb, a, cursor) == 1);
                }
                break;
            case FUNCTION: // function string
                break;
            default:
                a = getV(e);
                if (a < cursor) {
                    if (isNumber(a, cursor)) {
                        long l = toLong(a, cursor);
                        if (l <= Integer.MAX_VALUE && l >= Integer.MIN_VALUE) {
                            cell.setNv((int) l);
                        } else {
                            cell.setLv(l);
                        }
                    } else if (isDouble(a, cursor)) {
                        cell.setDv(toDouble(a, cursor));
                    } else {
                        cell.setSv(toString(a, cursor));
                    }
                }
        }

        // end of cell
        cursor = e;

        return cell;
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
        for (; a < b; ) {
            char c = cb[a++];
            if (c < '0' || c > '9') break;
        }
        return a == b;
    }

    /**
     * FIXME check double
     *
     * @param a
     * @param b
     * @return
     */
    private boolean isDouble(int a, int b) {
        if (a == b) return false;
        if (cb[a] == '-') a++;
        for (char i = 0; a < b; ) {
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
     *
     * @param e the last index in char buffer
     * @return the end index of string value
     */
    private int getT(int e) {
        for (; cursor < e && (cb[cursor] != '<' || cb[cursor + 1] != 't'
            || cb[cursor + 2] != '>'); cursor++) ;
        if (cursor == e) return cursor;
        int a = cursor += 3;
        for (; cursor < e && (cb[cursor] != '<' || cb[cursor + 1] != '/'
            || cb[cursor + 2] != 't' || cb[cursor + 3] != '>'); cursor++) ;
        return a;
    }

    /**
     * The string index in shared string table
     *
     * @param e the last index in char buffer
     * @return the end index of int value
     */
    private int getV(int e) {
        for (; cursor < e && (cb[cursor] != '<' || cb[cursor + 1] != 'v'
            || cb[cursor + 2] != '>'); cursor++) ;
        if (cursor == e) return cursor;
        int a = cursor += 3;
        for (; cursor < e && (cb[cursor] != '<' || cb[cursor + 1] != '/'
            || cb[cursor + 2] != 'v' || cb[cursor + 3] != '>'); cursor++) ;
        return a;
    }

    /**
     * function string
     *
     * @param e the last index in char buffer
     * @return the end index of function value
     */
    @SuppressWarnings("unused")
    private int getF(int e) {
        // undo
        // return end index of row
        return e;
    }

    /**
     * Convert to column index
     *
     * @param a the start index
     * @param b the end index
     * @return the cell index
     */
    private int toCellIndex(int a, int b) {
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

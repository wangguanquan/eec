/*
 * Copyright (c) 2019-2021, guanquan.wang@yandex.com All Rights Reserved.
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

import org.ttzero.excel.entity.style.Styles;

import java.util.Arrays;

import static org.ttzero.excel.reader.Cell.BOOL;
import static org.ttzero.excel.reader.Cell.NUMERIC;
import static org.ttzero.excel.reader.Cell.FUNCTION;
import static org.ttzero.excel.reader.Cell.SST;
import static org.ttzero.excel.reader.Cell.INLINESTR;
import static org.ttzero.excel.reader.Cell.BLANK;
import static org.ttzero.excel.reader.SharedStrings.toInt;
import static org.ttzero.excel.reader.SharedStrings.unescape;
import static org.ttzero.excel.util.StringUtil.swap;

/**
 * Row data, shared by the Row object in the same Sheet page.
 * The row data start column and the end column read the span
 * value. You can use the {@link #isEmpty()} method to
 * test whether the row data is an empty node. The empty node
 * is defined as: No values and styles and formats. Line like this.
 * <code><row r="x"/></code> You can get the data eq by cell
 * subscript like ResultSet: {@link #getInt(int)} to get the current
 * line The data in the second column, the subscript, starts at 0.
 * <p>
 * Create by guanquan.wang on 2018-09-22
 */
class XMLRow extends Row {
    private int startRow;
    private StringBuilder buf;
    private MergeCalc calcFun;
    private boolean hasCalc;

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

    @SuppressWarnings("unused")
    private XMLRow() { }

    XMLRow(SharedStrings sst, Styles styles, int startRow, MergeCalc calcFun) {
        this.sst = sst;
        this.styles = styles;
        this.startRow = startRow;
        buf = new StringBuilder();
        this.calcFun = calcFun;
        this.hasCalc = calcFun != null;
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
        int _f = from + 4, a; // skip '<row'
        for (; cb[_f] != '>' && _f < to; _f++) {
            if (cb[_f] <= ' ' && cb[_f + 1] == 'r' && cb[_f + 2] == '=') {
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
        int i = from + 4, _lc = lc;
        for (; cb[i] != '>'; i++) {
            if (cb[i] <= ' ' && cb[i + 1] == 's' && cb[i + 2] == 'p'
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
        if (hr != null && lc < hr.lc) lc = hr.lc;
        fc = fc - 1; // zero base
        if (cells == null || lc > cells.length) {
            cells = new Cell[lc > 0 ? lc : 100]; // default array length 100
        }
        // clear and share
        for (int n = 0, len = lc > 0 ? Math.max(lc, _lc) : cells.length; n < len; n++) {
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

        // Parse formula if exists and can parse
        if (hasCalc) {
            calcFun.accept(getRowNumber(), cells, !unknownLength ? lc - fc : -1);
        }

        // Parse cell value
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
            || cb[cursor + 2] > ' '); cursor++) ;
        // end of row
        if (cursor >= to) return null;
        cursor += 2;
        // find end of cell
        int e = cursor;
        for (; e < to && (cb[e] != '<' || cb[e + 1] != 'c' || cb[e + 2] > ' '); e++) ;

        Cell cell = null;
        // find type
        // n=numeric (default), s=string, b=boolean, str=function string
        char t = NUMERIC; // default
        int xf = 0, i = 1;
        for (; cb[cursor] != '>'; cursor++) {
            // Cell index
            if (cb[cursor] <= ' ' && cb[cursor + 1] == 'r' && cb[cursor + 2] == '=') {
                int a = cursor += 4;
                for (; cb[cursor] != '"'; cursor++) ;
                i = unknownLength ? (lc = toCellIndex(a, cursor)) : toCellIndex(a, cursor);
                cell = cells[i - 1];
            }
            // Cell type
            if (cb[cursor] <= ' ' && cb[cursor + 1] == 't' && cb[cursor + 2] == '=') {
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
            if (cb[cursor] <= ' ' && cb[cursor + 1] == 's' && cb[cursor + 2] == '=') {
                int a = cursor += 4;
                for (; cb[cursor] != '"'; cursor++) ;
                xf = toInt(cb, a, cursor);
            }
        }

        if (cell == null) return null;

        // The style index
        cell.xf = xf;

        // Ignore Formula string default
        if (hasCalc) {
            int a = getF(cell, e);
            // Inner text
            if (a < cursor) {
                cell.fv = unescape(buf, cb, a, cursor);
                // TODO header row is null
                if (hr != null && cell.si != -1)
                    hr.setCalc(cell.si, cell.fv);
            }
            // TODO header row is null
            // Function string is shared
            else if (hr != null && cell.si > -1) {
                // Get from ref
                cell.fv = hr.getCalc(cell.si, (getRowNumber() << 16) | i);
            }
        }
        // Get value
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
                a = getV(e);
                if (a == cursor) { // null value
                    cell.setT(BLANK); // Reset type to BLANK if null value
                } else {
                    cell.setSv(unescape(buf, cb, a, cursor));
                }
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
        return Double.parseDouble(toString(a, b));
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

    /* Found specify target  */
    private int get(int e, char c) {
        // Ignore all attributes
        return get(null, e, c, null);
    }

    /**
     * Parses the value and attributes of the specified tag
     *
     * @param cell current cell
     * @param e the last character index
     * @param c the specified tag
     * @param attrConsumer an attribute consumer
     * @return the start index of value
     */
    private int get(Cell cell, int e, char c, Attribute attrConsumer) {
        for (; cursor < e && (cb[cursor] != '<' || cb[cursor + 1] != c
            || cb[cursor + 2] != '>' && cb[cursor + 2] > ' ' && cb[cursor + 2] != '/'); cursor++) ;
        if (cursor == e) return cursor;

        int a;
        if (cb[cursor + 2] == '>') {
            a = cursor += 3;
        }
        // Some other attributes
        else if (cb[cursor + 2] == ' ') {
            int i = cursor + 3;
            for (; cursor < e && cb[cursor] != '>'; cursor++) ;
            // If parse attributes
            if (attrConsumer != null)
                attrConsumer.accept(cell, cb, i, cb[cursor - 1] != '/' ? cursor : cursor - 1);

            cursor++;
            if (cb[cursor - 2] == '/' || cursor == e) {
                return cursor;
            }
            a = cursor;
        }
        // Empty tag
        else if (cb[cursor + 2] == '/') {
            cursor += 3;
            return cursor;
        }
        else {
            a = cursor += 3;
        }

        for (; cursor < e && (cb[cursor] != '<' || cb[cursor + 1] != '/'
            || cb[cursor + 2] != c || cb[cursor + 3] != '>'); cursor++) ;
        return a;
    }

    /**
     * Found text tag range
     *
     * Code like this {@code <is><t>cell value</t></is>}
     *
     * @param e the last index in char buffer
     * @return the end index of string value
     */
    private int getT(int e) {
        return get(e, 't');
    }

    /**
     * Found value tag range
     *
     * Code like this {@code <v>0</v>
     *
     * @param e the last index in char buffer
     * @return the end index of int value
     */
    private int getV(int e) {
        return get(e, 'v');
    }

    /**
     * Found the Function tag range
     * Code like this {@code <f t="shared" ref="B1:B10" si="0">SUM(A1:A10)</f>
     *
     * @param e the last index in char buffer
     * @return the end index of function value
     */
    private int getF(Cell cell, int e) {
        return get(cell, e, 'f', this::parseFunAttr);
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

    /* Parse function tag's attribute */
    private void parseFunAttr(Cell cell, char[] cb, int a, int b) {
        // t="shared" ref="B2:B3" si="0"
        String[] values = new String[10];
        int index = 0;
        boolean sv = false; // is string value
        for (int i = a ; ; ) {
            for (; a < b && cb[a] > ' ' && cb[a] != '='; a++) ;
            values[index++] = new String(cb, i, sv ? a - i - 1 : a - i);
            sv = false;

            if (a + 1 < b) {
                // String value
                if (cb[a + 1] == '"') {
                    a += 2;
                    sv = true;
                }
                // Boolean value
                else if (cb[a + 1] == ' ') {
                    values[index++] = "1";
                    a++;
                }
                else a++;
                i = a;
            } else break;
        }

        if (index < 2 || (index & 1) == 1) {
            logger.warn("The function format error.[{}]", new String(cb, a, b - a));
            return;
        }

        // Sort like t, si, ref
        for (int i = 0, len = index >> 1; i < len; i++) {
            int _i = i << 1, vl = values[_i].length();
            if (vl - 1 == i) {
                continue;
            }
            // Will be sort
            int _n = vl - 1;
            if (_n > index - 1) {
                logger.warn("Unknown attribute on function tag.[{}]", values[_i]);
                return;
            }
            swap(values, _n << 1, _i);
            swap(values, (_n << 1) + 1, _i + 1);
        }

        int si = Integer.parseInt(values[3]);

        // Append and share href
        if (index > 4) {
            hr.addRef(si, values[5], null);
        }

        // Storage formula shared id
        cell.si = si;
    }

    /**
     * Attribute consumer
     */
    @FunctionalInterface
    private interface Attribute {
        /**
         * Performs this operation on the given argument.
         *
         * @param cell current cell
         * @param cb characters for the entire attribute
         * @param a start index
         * @param b end index
         */
        void accept(Cell cell, char[] cb, int a, int b);
    }
}

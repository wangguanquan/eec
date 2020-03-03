/*
 * Copyright (c) 2017-2018, guanquan.wang@yandex.com All Rights Reserved.
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

import static org.ttzero.excel.reader.Cell.BOOL;
import static org.ttzero.excel.reader.Cell.EMPTY_TAG;
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
 *
 * @author guanquan.wang on 2018-09-22
 */
class XMLRow extends Row {
    int startRow;
    StringBuilder buf;

    /**
     * The number of row. (one base)
     *
     * @return int value
     */
    @Override
    public int getRowNumber() {
        if (index == -1)
            searchRowNumber();
        // The first row index is one
        return index;
    }

    @SuppressWarnings("unused")
    XMLRow() { }

    XMLRow(SharedStrings sst, Styles styles, int startRow) {
        this.sst = sst;
        this.styles = styles;
        this.startRow = startRow;
        buf = new StringBuilder();
    }

    /////////////////////////unsafe////////////////////////
    char[] cb;
    int from, to;
    int cursor, e;

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

    private void searchRowNumber() {
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
    }

    int searchSpan() {
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
            else cells[n] = new Cell((short) (n + 1));
        }
        return i;
    }

    /**
     * Loop parse cell
     */
    void parseCells() {
        cursor = searchSpan();
        for (; cb[cursor++] != '>'; ) ;
        unknownLength = lc < 0;

        Cell cell;
        int index = 0;
        // Parse cell value
        if (unknownLength) {
            for(; (cell = nextCell()) != null; index++, parseCellValue(cell));
        } else {
            for(; index < lc && (cell = nextCell()) != null; parseCellValue(cell)) ;
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
        e = cursor;
        for (; e < to && (cb[e] != '<' || cb[e + 1] != 'c' || cb[e + 2] > ' '); e++) ;

        Cell cell = null;
        // find type
        // n=numeric (default), s=string, b=boolean, str=function string
        char t = NUMERIC; // default
        int xf = 0, i;
        for (; cb[cursor] != '>'; cursor++) {
            // Cell index
            if (cb[cursor] <= ' ' && cb[cursor + 1] == 'r' && cb[cursor + 2] == '=') {
                int a = cursor += 4;
                for (; cb[cursor] != '"'; cursor++) ;
                i = unknownLength ? (lc = toCellIndex(cb, a, cursor)) : toCellIndex(cb, a, cursor);
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
        cell.t = t;

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
        for (char i = 0, e = 0; a < b; ) {
            char c = cb[a++];
            if (i > 1 || e > 1) return false;
            if (c == '.') i++;
            else if (c == 'e' || c == 'E') e++;
            else if (c < '0' || c > '9') return false;
        }
        return true;
    }

    /* Found specify target  */
    private int get(char c) {
        // Ignore all attributes
        return get(null, c, null);
    }

    /**
     * Parses the value and attributes of the specified tag
     *
     * @param cell current cell
     * @param c the specified tag
     * @param attrConsumer an attribute consumer
     * @return the start index of value
     */
    int get(Cell cell, char c, Attribute attrConsumer) {
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

        // Found end tag
        for (; cursor < e && (cb[cursor] != '<' || cb[cursor + 1] != '/'
            || cb[cursor + 2] != c || cb[cursor + 3] != '>'); cursor++) ;

        return a;
    }

    /**
     * Found text tag range
     *
     * Code like this {@code <is><t>cell value</t></is>}
     *
     * @return the end index of string value
     */
    private int getT() {
        return get('t');
    }

    /**
     * Found value tag range
     *
     * Code like this {@code <v>0</v>
     *
     * @return the end index of int value
     */
    private int getV() {
        return get('v');
    }

    /**
     * Parse cell value
     *
     * @param cell current {@link Cell}
     */
    void parseCellValue(Cell cell) {
        // @Mark: Ignore Formula string default

        // Get value
        int a;
        switch (cell.t) {
            case INLINESTR: // inner string
                a = getT();
                if (a < cursor) {
                    cell.setSv(unescape(buf, cb, a, cursor));
                } else { // null value
                    cell.setT(BLANK); // Reset type to BLANK if null value
                }
                break;
            case SST: // shared string lazy get
                a = getV();
                cell.setNv(toInt(cb, a, cursor));
                cell.setT(SST);
                break;
            case BOOL: // boolean value
                a = getV();
                if (cursor - a == 1) {
                    cell.setBv(toInt(cb, a, cursor) == 1);
                } else cell.setBv(false);
                break;
            case FUNCTION: // function string
                a = getV();
                if (a < cursor) {
                    cell.setSv(unescape(buf, cb, a, cursor));
                } else { // null value
                    cell.setT(BLANK); // Reset type to BLANK if null value
                }
                break;
            default:
                a = getV();
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
                // Maybe the cell should be merged
                else cell.setT(EMPTY_TAG);
        }

        // end of cell
        cursor = e;
    }

    XMLCalcRow asCalcRow() {
        return !(this instanceof XMLCalcRow) ? new XMLCalcRow(this) : (XMLCalcRow) this;
    }

    XMLMergeRow asMergeRow() {
        return !(this instanceof XMLMergeRow) ? new XMLMergeRow(this) : (XMLMergeRow) this;
    }

    /**
     * Attribute consumer
     */
    @FunctionalInterface
    interface Attribute {
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

/**
 * Cell with Calc
 */
class XMLCalcRow extends XMLRow {
    private MergeCalcFunc calcFun;

    XMLCalcRow(SharedStrings sst, Styles styles, int startRow, MergeCalcFunc calcFun) {
        this.sst = sst;
        this.styles = styles;
        this.startRow = startRow;
        this.buf = new StringBuilder();
        this.calcFun = calcFun;
    }

    XMLCalcRow(XMLRow row) {
        this.sst = row.sst;
        this.styles = row.styles;
        this.startRow = row.startRow;
        this.buf = row.buf;
    }

    XMLCalcRow setCalcFun(MergeCalcFunc calcFun) {
        this.calcFun = calcFun;
        return this;
    }

    /**
     * Loop parse cell
     */
    void parseCells() {
        int index = 0;
        cursor = searchSpan();
        for (; cb[cursor++] != '>'; ) ;
        unknownLength = lc < 0;

        // Parse formula if exists and can parse
        calcFun.accept(getRowNumber(), cells, !unknownLength ? lc - fc : -1);

        Cell cell;
        // Parse cell value
        if (unknownLength) {
            for(; (cell = nextCell()) != null; index++, parseCellValue(cell));
        } else {
            for(; index < lc && (cell = nextCell()) != null; parseCellValue(cell)) ;
        }
    }

    /**
     * Parse cell value
     *
     * @param cell current {@link Cell}
     */
    @Override
    void parseCellValue(Cell cell) {
        // Parse calc
        parseCalcFunc(cell);

        // Parse value
        super.parseCellValue(cell);
    }

    /**
     * Parse calc on reader cells
     *
     * @param cell current {@link Cell}
     */
    private void parseCalcFunc(Cell cell) {
        int a = getF(cell);
        // Inner text
        if (a < cursor) {
            cell.fv = unescape(buf, cb, a, cursor);
            if (cell.si > -1) setCalc(cell.si, cell.fv);
        }
        // Function string is shared
        else if (cell.si > -1) {
            // Get from ref
            cell.fv = getCalc(cell.si, (getRowNumber() << 14) | cell.i);
        }
    }


    /**
     * Found the Function tag range
     * Code like this {@code <f t="shared" ref="B1:B10" si="0">SUM(A1:A10)</f>
     *
     * @param cell current {@link Cell}
     * @return the end index of function value
     */
    private int getF(Cell cell) {
        return get(cell, 'f', this::parseFunAttr);
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
            addRef(si, values[5]);
        }

        // Storage formula shared id
        cell.si = si;
    }
}

/**
 * Copy value on merge cells
 */
class XMLMergeRow extends XMLRow {
    // InterfaceFunction
    private MergeValueFunc func;

    XMLMergeRow(XMLRow row) {
        this.sst = row.sst;
        this.styles = row.styles;
        this.startRow = row.startRow;
        this.buf = row.buf;
    }

    XMLMergeRow setCopyValueFunc(MergeValueFunc func) {
        this.func = func;
        return this;
    }

    /**
     * Parse cell value
     *
     * @param cell current {@link Cell}
     */
    @Override
    void parseCellValue(Cell cell) {

        // Parse value
        super.parseCellValue(cell);

        // Setting/copy value if merged
        func.accept(getRowNumber(), cell);
    }
}

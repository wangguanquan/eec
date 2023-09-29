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

import org.ttzero.excel.entity.TooManyColumnsException;
import org.ttzero.excel.entity.style.Styles;
import org.ttzero.excel.manager.Const;

import java.math.BigDecimal;

import static org.ttzero.excel.reader.Cell.BOOL;
import static org.ttzero.excel.reader.Cell.NUMERIC;
import static org.ttzero.excel.reader.Cell.FUNCTION;
import static org.ttzero.excel.reader.Cell.SST;
import static org.ttzero.excel.reader.Cell.INLINESTR;
import static org.ttzero.excel.reader.Cell.UNALLOCATED;
import static org.ttzero.excel.reader.SharedStrings.toInt;
import static org.ttzero.excel.reader.SharedStrings.escape;
import static org.ttzero.excel.util.StringUtil.swap;

/**
 * Row data, shared by the Row object in the same Sheet page.
 * The row data start column and the end column read the span
 * value. You can use the {@link #isEmpty()} method to
 * test whether the row data is an empty node. The empty node
 * is defined as: No values and styles and formats. Line like this.
 * {@code <row r="x"/>} You can get the data eq by cell
 * subscript like ResultSet: {@link #getInt(int)} to get the current
 * line The data in the second column, the subscript, starts at 0.
 *
 * @author guanquan.wang on 2018-09-22
 */
public class XMLRow extends Row {
    protected int startRow;

    /**
     * The number of row. (one base)
     *
     * @return int value
     */
    @Override
    public int getRowNum() {
        if (index == -1)
            searchRowNum();
        // The first row index is one
        return index;
    }

    @SuppressWarnings("unused")
    public XMLRow() {
        this.startRow = 1;
    }

    public XMLRow(SharedStrings sst, Styles styles, int startRow) {
        init(sst, styles, startRow);
    }

    public XMLRow init(SharedStrings sst, Styles styles, int startRow) {
        this.sst = sst;
        this.styles = styles;
        this.startRow = startRow;
        return this;
    }

    /////////////////////////unsafe////////////////////////
    protected char[] cb;
    protected int from, to;
    protected int cursor, e;

    ///////////////////////////////////////////////////////
    protected XMLRow with(char[] cb, int from, int size) {
//        LOGGER.debug(new String(cb, from, size));
        this.cb = cb;
        this.from = from;
        this.to = from + size;
        this.cursor = from;
        this.index = this.lc = -1;
        parseCells();
        return this;
    }

    /* empty row*/
    protected XMLRow empty(char[] cb, int from, int size) {
//        LOGGER.debug(new String(cb, from, size));
        this.cb = cb;
        this.from = from;
        this.to = from + size;
        this.cursor = from;
        this.index = -1;
        this.fc = this.lc = -1;
        return this;
    }

    private void searchRowNum() {
        if (from >= to || cb == null) return;
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

    protected int searchSpan() {
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
        if (hr != null && lc < hr.lc) lc = hr.lc;
        if (fc <= 0 || fc >= lc) fc = this.startRow;
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
    protected void parseCells() {
        cursor = searchSpan();
        for (; cb[cursor++] != '>'; ) ;
        unknownLength = lc < 0;

        // Parse cell value
        for (Cell cell; (cell = nextCell()) != null; parseCellValue(cell)) ;
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
        int xf = 0, i = 0;
        for (; cb[cursor] != '>'; cursor++) {
            // Cell index
            if (cb[cursor] <= ' ' && cb[cursor + 1] == 'r' && cb[cursor + 2] == '=') {
                int a = cursor += 4;
                for (; cb[cursor] != '"'; cursor++) ;
                i = toCellIndex(cb, a, cursor);
                // The `spans` attribute is not be set
                if (i - 1 >= cells.length) {
                    // Bound check
                    if (i - 1 > Const.Limit.MAX_COLUMNS_ON_SHEET) {
                        throw new TooManyColumnsException(i, Const.Limit.MAX_COLUMNS_ON_SHEET);
                    }
                    // Resize cell buffer
                    cells = copyCells(Math.min(i + 99, Const.Limit.MAX_COLUMNS_ON_SHEET));
                }
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
                // -> Other unknown case
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
        if (lc < i) lc = i;

        return cell;
    }

    protected static long toLong(char[] cb, int a, int b) {
        boolean _n;
        if (_n = cb[a] == '-') a++;
        long n = cb[a++] - '0';
        for (; b > a; ) {
            n = n * 10 + cb[a++] - '0';
        }
        return _n ? -n : n;
    }

//    protected static double toDouble(char[] cb, int a, int b) {
//        return Double.parseDouble(new String(cb, a, b - a));
//    }


    /* Found specify target  */
    protected int get(char c) {
        for (; cursor < e && (cb[cursor] != '<' || cb[cursor + 1] != c
            || cb[cursor + 2] != '>' && cb[cursor + 2] > ' ' && cb[cursor + 2] != '/'); cursor++) ;
        if (cursor == e) return cursor;

        int a;
        if (cb[cursor + 2] == '>') a = cursor += 3;

        // Some other attributes
        else if (cb[cursor + 2] == ' ') {
            for (; cursor < e && cb[cursor] != '>'; cursor++) ;

            cursor++;
            if (cb[cursor - 2] == '/' || cursor == e) return cursor;
            a = cursor;
        }
        // Empty tag
        else if (cb[cursor + 2] == '/') {
            cursor += 3;
            return cursor;
        } else a = cursor += 3;

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
    protected int getT() {
        return get('t');
    }

    /**
     * Found value tag range
     *
     * Code like this {@code <v>0</v>}
     *
     * @return the end index of int value
     */
    protected int getV() {
        return get('v');
    }

    /**
     * Parse cell value
     *
     * @param cell current {@link Cell}
     */
    protected void parseCellValue(Cell cell) {
        // @Mark: Ignore Formula string default

        // Get value
        int a;
        switch (cell.t) {
            case INLINESTR: // inner string
                a = getT();
                if (a < cursor) {
                    cell.setSv(escape(cb, a, cursor));
                } else { // null value
                    cell.blank(); // Reset type to BLANK if null value
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
                    cell.setSv(escape(cb, a, cursor));
                } else { // null value
                    cell.blank(); // Reset type to BLANK if null value
                }
                break;
            case UNALLOCATED: return; // Not [break] header
            default:
                a = getV();
                if (a < cursor) {
                    int t = testNumberType(cb, a, cursor);
                    // -1: not a number
                    // 0: empty
                    // 1: int
                    // 2: long
                    // 3: double
                    switch (t) {
                        case 3: cell.setMv(new BigDecimal(cb, a, cursor - a)); break;
                        case 2: {
                            long l = toLong(cb, a, cursor);
                            if (l > Integer.MAX_VALUE || l < Integer.MIN_VALUE) cell.setLv(l);
                            else cell.setNv((int) l);
                            break;
                        }
                        case 1: cell.setNv(toInt(cb, a, cursor));    break;
                        case 0: cell.emptyTag();                     break;
                        default: cell.setSv(escape(cb, a, cursor));
                    }
                }
                // Maybe the cell should be merged
                else cell.emptyTag();
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

}

/**
 * Cell with Calc
 */
class XMLCalcRow extends XMLRow {
    private MergeCalcFunc calcFun;
    private boolean hasCalcFunc;

    XMLCalcRow(SharedStrings sst, Styles styles, int startRow, MergeCalcFunc calcFun) {
        this.sst = sst;
        this.styles = styles;
        this.startRow = startRow;
        this.calcFun = calcFun;
        this.hasCalcFunc = calcFun != null;
    }

    XMLCalcRow(XMLRow row) {
        this.sst = row.sst;
        this.styles = row.styles;
        this.startRow = row.startRow;
    }

    XMLCalcRow setCalcFun(MergeCalcFunc calcFun) {
        this.calcFun = calcFun;
        hasCalcFunc = calcFun != null;
        return this;
    }

    /**
     * Loop parse cell
     */
    @Override
    protected void parseCells() {
        cursor = searchSpan();
        for (; cb[cursor++] != '>'; ) ;
        unknownLength = lc < 0;

        // Parse formula if exists and can parse
        if (hasCalcFunc) {
            calcFun.accept(getRowNum(), cells, !unknownLength ? lc - fc : -1);
        }

        // Parse cell value
        for (Cell cell; (cell = nextCell()) != null; parseCellValue(cell)) ;
    }

    /**
     * Parse cell value
     *
     * @param cell current {@link Cell}
     */
    @Override
    protected void parseCellValue(Cell cell) {
        // If cell has formula
        if (cell.f || !hasCalcFunc) {
            // Parse calc
            parseCalcFunc(cell);
        }

        // Parse value
        super.parseCellValue(cell);
    }

    /**
     * Parse calc on reader cells
     *
     * @param cell current {@link Cell}
     */
    private void parseCalcFunc(Cell cell) {
        int _cursor = cursor, a = getF(cell);
        // Reset the formula flag
        cell.f = a < cursor;
        // Tag <f> Not Found
        if (a == cursor) {
            cursor = _cursor;
            return;
        }
        // Inner text
        if (a < cursor) {
            cell.fv = escape(cb, a, cursor);
            if (cell.si > -1) setCalc(cell.si, cell.fv);
        }
        // Function string is shared
        else if (cell.si > -1) {
            // Get from ref
            cell.fv = getCalc(cell.si, (getRowNum() << 14) | cell.i);
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
        for (; cursor < e && (cb[cursor] != '<' || cb[cursor + 1] != 'f'
            || cb[cursor + 2] != '>' && cb[cursor + 2] > ' ' && cb[cursor + 2] != '/'); cursor++) ;
        if (cursor == e) return cursor;

        int a;
        if (cb[cursor + 2] == '>') a = cursor += 3;

        // Some other attributes
        else if (cb[cursor + 2] == ' ') {
            int i = cursor + 3;
            for (; cursor < e && cb[cursor] != '>'; cursor++) ;
            // If parse attributes
            parseFunAttr(cell, cb, i, cb[cursor - 1] != '/' ? cursor : cursor - 1);

            cursor++;
            if (cb[cursor - 2] == '/' || cursor == e) return cursor;
            a = cursor;
        }
        // Empty tag
        else if (cb[cursor + 2] == '/') {
            cursor += 3;
            return cursor;
        } else a = cursor += 3;

        // Found end tag
        for (; cursor < e && (cb[cursor] != '<' || cb[cursor + 1] != '/'
            || cb[cursor + 2] != 'f' || cb[cursor + 3] != '>'); cursor++) ;

        return a;
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
            LOGGER.warn("The function format error.[{}]", new String(cb, a, b - a));
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
                LOGGER.warn("Unknown attribute on function tag.[{}]", values[_i]);
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
    protected MergeValueFunc func;
    // A merge cells grid
    protected Grid mergeCells;

    XMLMergeRow(XMLRow row) {
        this.sst = row.sst;
        this.styles = row.styles;
        this.startRow = row.startRow;
    }

    XMLMergeRow setCopyValueFunc(Grid mergeCells, MergeValueFunc func) {
        this.mergeCells = mergeCells;
        this.func = func;
        return this;
    }


    /**
     * Loop parse cell
     */
    @Override
    protected void parseCells() {
        cursor = searchSpan();
        for (; cb[cursor++] != '>'; ) ;
        unknownLength = lc < 0;

        // Parse cell value
        int i = 1, r = getRowNum();
        for (Cell cell; (cell = nextCell()) != null; ) {
            if (cell.i > i) for (; i < cell.i; i++) parseCellValue(cells[i - 1]);
            parseCellValue(cell);
            i++;
        }

        /*
         Some tools handle merged cells that ignore all cells
          in the merged range except for the first one,
          so compatibility is required here for cells that are outside the spans range
         */
        for (; mergeCells.test(r, i); i++) {
            if (lc < i) {
                // Give a new cells
                if (cells.length < i) cells = copyCells(i);
                lc = i;
            }
            parseCellValue(cells[i - 1]);
        }
    }


    /**
     * Parse cell value
     *
     * @param cell current {@link Cell}
     */
    @Override
    protected void parseCellValue(Cell cell) {

        // Parse value
        super.parseCellValue(cell);

        // Setting/copy value if merged
        func.accept(getRowNum(), cell);
    }
}

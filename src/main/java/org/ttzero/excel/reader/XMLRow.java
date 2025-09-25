/*
 * Copyright (c) 2017-2018, guanquan.wang@hotmail.com All Rights Reserved.
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
import org.ttzero.excel.util.StringUtil;

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
    /**
     * @deprecated 未使用，下个版本删除
     */
    @Deprecated
    protected int startRow;

    /**
     * The number of row. (one base)
     *
     * @return int value
     */
    @Override
    public int getRowNum() {
        if (rowNum == -1) rowNum = index = searchRowNum();
        return rowNum;
    }

    public XMLRow() { }

    public XMLRow(SharedStrings sst, Styles styles) {
        init(sst, styles);
    }

    public XMLRow init(SharedStrings sst, Styles styles) {
        this.sst = sst;
        this.styles = styles;
        return this;
    }

    @Deprecated
    public XMLRow(SharedStrings sst, Styles styles, int startRow) {
        init(sst, styles);
    }

    @Deprecated
    public XMLRow init(SharedStrings sst, Styles styles, int startRow) {
        return init(sst, styles);
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
        this.rowNum = this.index = this.lc = -1; // 兼容处理，后续删除index
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
        this.rowNum = this.index = -1; // 兼容处理，后续删除index
        this.fc = this.lc = -1;
        return this;
    }

    protected int searchRowNum() {
        if (from >= to || cb == null) return -1;
        int _f = from + 4, a; // skip '<row'
        for (; cb[_f] != '>' && _f < to; _f++) {
            if (cb[_f] <= ' ' && cb[_f + 1] == 'r' && cb[_f + 2] == '=') {
                a = _f += 4;
                for (; cb[_f] != '"' && _f < to; _f++) ;
                if (_f > a) {
                    return toInt(cb, a, _f); // 兼容处理，后续删除index
                }
                break;
            }
        }
        return -1;
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
                if (i > j) {
                    for (b = i - 1; cb[b] != ':' && b > j; b--) ;
                    if (++b < i) {
                        lc = toInt(cb, b, i);
                    }
                    if (j < --b) {
                        fc = toInt(cb, j, b);
                    }
                }
            }
        }
        if (hr != null && lc < hr.lc) lc = hr.lc;
        if (fc <= 0 || fc >= lc) fc = 1;
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
                } else if (n == 9 && cb[a] == 'i' && cb[a + 1] == 'n'
                    && cb[a + 2] == 'l' && cb[a + 6] == 'S' && cb[a + 8] == 'r') {
                    t = INLINESTR; // inlineStr
                } else if (n == 3 && cb[a] == 's' && cb[a + 1] == 't' && cb[a + 2] == 'r') {
                    t = FUNCTION; // function string
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
//        if (lc < i)
        lc = i;

        return cell;
    }

    protected static long toLong(char[] cb, int a, int b) {
        boolean _n = cb[a] == '-';
        if (_n) a++;
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
     * <p>Code like this {@code <is><t>cell value</t></is>}</p>
     *
     * @return the end index of string value
     */
    protected int getT() {
        return get('t');
    }

    /**
     * Found value tag range
     *
     * <p>Code like this {@code <v>0</v>}</p>
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
                    cell.setString(escape(cb, a, cursor));
                } else { // null value
                    cell.blank(); // Reset type to BLANK if null value
                }
                break;
            case SST: // shared string lazy get
                a = getV();
                cell.setInt(toInt(cb, a, cursor));
                cell.setT(SST);
                break;
            case BOOL: // boolean value
                a = getV();
                if (cursor - a == 1) {
                    cell.setBool(toInt(cb, a, cursor) == 1);
                } else cell.setBool(false);
                break;
            case FUNCTION: // function string
                a = getV();
                if (a < cursor) {
                    cell.setString(escape(cb, a, cursor));
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
                        case 3: cell.setDecimal(new BigDecimal(cb, a, cursor - a)); break;
                        case 2: {
                            long l = toLong(cb, a, cursor);
                            if (l > Integer.MAX_VALUE || l < Integer.MIN_VALUE) cell.setLong(l);
                            else cell.setInt((int) l);
                            break;
                        }
                        case 1: cell.setInt(toInt(cb, a, cursor));    break;
                        case 0: cell.emptyTag();                     break;
                        default: cell.setString(escape(cb, a, cursor));
                    }
                }
                // Maybe the cell should be merged
                else cell.emptyTag();
        }

        // end of cell
        cursor = e;
    }

//    @Deprecated
//    XMLCalcRow asCalcRow() {
//        return !(this instanceof XMLCalcRow) ? new XMLCalcRow(this) : (XMLCalcRow) this;
//    }
//
//    @Deprecated
//    XMLMergeRow asMergeRow() {
//        return !(this instanceof XMLMergeRow) ? new XMLMergeRow(this) : (XMLMergeRow) this;
//    }

    XMLFullRow asFullRow() {
        return this.getClass() != XMLFullRow.class ? new XMLFullRow(this) : (XMLFullRow) this;
    }
}

///**
// * Cell with Calc
// * @deprecated 使用 {@link XMLFullRow}代替,即将删除
// */
//@Deprecated
//class XMLCalcRow extends XMLFullRow {
//
//    XMLCalcRow(XMLRow row) {
//        super(row);
//    }
//
//    /**
//     * Loop parse cell
//     */
//    @Override
//    protected void parseCells() {
//        cursor = searchSpan();
//        for (; cb[cursor++] != '>'; ) ;
//        unknownLength = lc < 0;
//
//        // Parse formula if exists and can parse
//        if (hasCalcFunc) {
//            calcFun.accept(getRowNum(), cells, !unknownLength ? lc - fc : -1);
//        }
//
//        // Parse cell value
//        for (Cell cell; (cell = nextCell()) != null; subParseCellValue(cell)) ;
//    }
//
//    /**
//     * Parse cell value
//     *
//     * @param cell current {@link Cell}
//     */
//    @Override
//    protected void subParseCellValue(Cell cell) {
//        // Parse calc
//        parseCalcFunc(cell);
//
//        // Parse value
//        super.parseCellValue(cell);
//    }
//
//}
//
///**
// * Copy value on merge cells
// * @deprecated 使用 {@link XMLFullRow}代替,即将删除
// */
//@Deprecated
//class XMLMergeRow extends XMLFullRow {
//
//    XMLMergeRow(XMLRow row) {
//        super(row);
//    }
//
//    /**
//     * Loop parse cell
//     */
//    @Override
//    protected void parseCells() {
//        cursor = searchSpan();
//        for (; cb[cursor++] != '>'; ) ;
//        unknownLength = lc < 0;
//
//        // Parse cell value
//        int i = 1, r = getRowNum();
//        for (Cell cell; (cell = nextCell()) != null; ) {
//            if (cell.i > i) for (; i < cell.i; i++) subParseCellValue(cells[i - 1]);
//            subParseCellValue(cell);
//            i++;
//        }
//
//        /*
//         Some tools handle merged cells that ignore all cells
//          in the merged range except for the first one,
//          so compatibility is required here for cells that are outside the spans range
//         */
//        for (; mergeCells.test(r, i); i++) {
//            if (lc < i) {
//                // Give a new cells
//                if (cells.length < i) cells = copyCells(i);
//                lc = i;
//            }
//            subParseCellValue(cells[i - 1]);
//        }
//    }
//
//
//    /**
//     * Parse cell value
//     *
//     * @param cell current {@link Cell}
//     */
//    @Override
//    protected void subParseCellValue(Cell cell) {
//
//        // Parse value
//        super.parseCellValue(cell);
//
//        // Setting/copy value if merged
//        mergedFunc.accept(getRowNum(), cell);
//    }
//}

/**
 * Copy value on merge cells
 */
class XMLFullRow extends XMLRow {
    MergeCalcFunc calcFun;
    boolean hasCalcFunc;
    // InterfaceFunction
    MergeValueFunc mergedFunc;
    // A merge cells grid
    Grid mergeGrid;
    // height，只有当customHeight为1时height才会有值
    Double height;
    // 是否隐藏
    boolean hidden;

    XMLFullRow(XMLRow row) {
        this.sst = row.sst;
        this.styles = row.styles;
    }

    @Override
    protected XMLFullRow empty(char[] cb, int from, int size) {
        super.empty(cb, from, size);
        searchSpan0(); // 解析行高
        return this;
    }

    XMLFullRow setCalcFun(MergeCalcFunc calcFun) {
        this.calcFun = calcFun;
        hasCalcFunc = calcFun != null;
        return this;
    }

    /**
     * Loop parse cell
     */
    @Override
    protected void parseCells() {
        height = null; // 重置行高
        hidden = false; // 重置隐藏
        cursor = searchSpan0();
        for (; cb[cursor++] != '>'; ) ;
        unknownLength = lc < 0;

        // Parse formula if exists and can parse
        if (hasCalcFunc) {
            calcFun.accept(getRowNum(), cells, lc > fc ? lc - fc : -1);
        }

        // Parse cell value
        int i = 1, r = getRowNum();
        for (Cell cell; (cell = nextCell()) != null; ) {
            if (cell.i > i) for (; i < cell.i; i++) subParseCellValue(cells[i - 1]);
            subParseCellValue(cell);
            i++;
        }

        /*
         Some tools handle merged cells that ignore all cells
          in the merged range except for the first one,
          so compatibility is required here for cells that are outside the spans range
         */
        for (; mergeGrid.test(r, i); i++) {
            if (lc < i) {
                // Give a new cells
                if (cells.length < i) cells = copyCells(i);
                lc = i;
            }
            subParseCellValue(cells[i - 1]);
        }
    }

    /**
     * Parse cell value
     *
     * @param cell current {@link Cell}
     */
    protected void subParseCellValue(Cell cell) {
        // Parse calc
        parseCalcFunc(cell);

        // Parse value
        super.parseCellValue(cell);

        // Setting/copy value if merged
        mergedFunc.accept(getRowNum(), cell);
    }

    /**
     * Parse calc on reader cells
     *
     * @param cell current {@link Cell}
     */
    void parseCalcFunc(Cell cell) {
        if (cell.t == UNALLOCATED) return;
        int _cursor = cursor, a = getF(cell);
        // Reset the formula flag
        cell.f = a < cursor || cell.si >= 0;
        // Tag <f> Not Found
        if (a == e) {
            cursor = _cursor;
            return;
        }
        // Inner text
        if (a < cursor) {
            cell.formula = escape(cb, a, cursor);
            if (cell.si > -1) setCalc(cell.si, cell.formula);
        }
        // Function string is shared
        else if (cell.si > -1) {
            // Get from ref
            cell.formula = getCalc(cell.si, (((long) getRowNum()) << 14) | cell.i);
        }
    }


    /**
     * Found the Function tag range
     * <p>Code like this {@code <f t="shared" ref="B1:B10" si="0">SUM(A1:A10)</f></p>
     *
     * @param cell current {@link Cell}
     * @return the end index of function value
     */
    int getF(Cell cell) {
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
    void parseFunAttr(Cell cell, char[] cb, int a, int b) {
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

        if ("si".equals(values[2]) && StringUtil.isNotEmpty(values[3])) {
            int si = Integer.parseInt(values[3]);

            // Append and share href
            if (index > 4) {
                addRef(si, values[5]);
            }

            // Storage formula shared id
            cell.si = si;
        }
    }

    XMLFullRow setCopyValueFunc(Grid mergeGrid, MergeValueFunc mergedFunc) {
        this.mergeGrid = mergeGrid;
        this.mergedFunc = mergedFunc;
        return this;
    }

    /**
     * 解析spans和行高
     *
     * @return 游标
     */
    int searchSpan0() {
        int idx = super.searchSpan(), i = from + 4;
        Double ht = null;
        for (; cb[i] != '>'; i++) {
            // 查找ht属性
            if (cb[i] <= ' ' && cb[i + 1] == 'h' && cb[i + 2] == 't' && (cb[i + 3] == '=' || cb[i + 3] <= ' ')) {
                i += 5;
                int j = i;
                for (; cb[i] != '"' && cb[i] != '>'; i++) ;
                if (i > j && cb[i] == '"') ht = Double.valueOf(new String(cb, j, i - j).trim());
            }
//            else if (cb[i] <= ' ' && cb[i + 1] == 'c' && cb[i + 2] == 'u' && cb[i + 3] == 's'
//                && cb[i + 4] == 't' && cb[i + 5] == 'o' && cb[i + 6] == 'm' && cb[i + 7] == 'H'
//                && cb[i + 8] == 'e' && cb[i + 9] == 'i' && cb[i + 10] == 'g' && cb[i + 11] == 'h'
//                && cb[i + 12] == 't' && (cb[i + 13] == '=' || cb[i + 13] <= ' ')) {
//                i += 15;
//                if (cb[i] == '1') cht = 1;
//            }
            // 查找hidden属性
            else if (cb[i] <= ' ' && cb[i + 1] == 'h' && cb[i + 2] == 'i' && cb[i + 3] == 'd'
                && cb[i + 4] == 'd' && cb[i + 5] == 'e' && cb[i + 6] == 'n' && (cb[i + 7] == '=' || cb[i + 7] <= ' ')) {
                i += 9;
                if (cb[i] == '1') hidden = true;
            }
        }
//        if (cht == 1 && ht != null) height = ht;
        if (ht != null) height = ht;
        return Math.max(i, idx);
    }

    @Override
    public Double getHeight() {
        return height;
    }

    @Override
    public boolean isHidden() {
        return hidden;
    }
}

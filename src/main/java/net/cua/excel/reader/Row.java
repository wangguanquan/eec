package net.cua.excel.reader;

import net.cua.excel.annotation.DisplayName;
import net.cua.excel.util.DateUtil;
import net.cua.excel.util.StringUtil;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.math.BigDecimal;
import java.sql.Timestamp;
import java.util.Arrays;
import java.util.Date;
import java.util.StringJoiner;

/**
 *
 * Create by guanquan.wang at 2018-09-22
 */
public class Row {
    Logger logger = LogManager.getLogger(getClass());
    private int rowNumber = -1, span = -1;
    private Cell[] cells;
    private SharedString sst;
    private HeaderRow hr;

    public int getRowNumber() {
        if (rowNumber == -1)
            searchRowNumber();
        return rowNumber;
    }

    private Row(){}

    Row(SharedString sst) {
        this.sst = sst;
    }

    ///////////////////////////safe///////////////////////
    Row safeWith(char[] cb, int from, int size) {
        if (this.cb == null || this.cb.length < cb.length) {
            this.cb = new char[size];
        }
        System.arraycopy(cb, from, this.cb, 0, size);
        this.from = 0;
        this.to = size;
        this.cursor = 0;
        this.rowNumber = this.span = -1;
        parseCells();
        return this;
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
        this.rowNumber = this.span = -1;
        parseCells();
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
                int b;i += 8;
                for (; cb[i] != '"' && cb[i] != '>'; i++);
                for (b = i - 1; cb[b] != ':'; b--);
                if (++b < i) {
                    span = toInt(b, i);
                }
            }
        }
        if (cells == null || span > cells.length) {
            cells = new Cell[span > 0 ? span : 100]; // default array length 100
        }
        // clear and share
        for (int n = 0; n < span; n++) {
            if (cells[n] != null) cells[n].clear();
            else cells[n] = new Cell();
        }
        return i;
    }

    private void parseCells() {
        int index = 0;
        cursor = searchSpan();
        for (; cb[cursor++] != '>'; );
        while (index < span && nextCell() != null);
    }

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
                cell = cells[toCellIndex(a, cursor) - 1];
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
                    cell.setSv(toString(a, cursor));
                }
                cell.setT('s'); // reset type to string
                break;
            case 's': // shared string
                a = getV(e);
                cell.setSv(sst.get(toInt(a, cursor)));
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

    private boolean isDouble(int a, int b) {
        if (a == b) return false;
        if (cb[a] == '-') a++;
        for (char i = 0 ; a < b; ) {
            char c = cb[a++];
            if (i > 1) return false;
            if (c >= '0' && c <= '9') continue;
            if (c == '.') c++;
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

    private int getF(int e) {
        // undo
        // return end index of row
        return e;
    }

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

    @Override public String toString() {
        StringJoiner joiner = new StringJoiner(" | ");
        // show row number
//        joiner.add(String.valueOf(getRowNumber()));
        for (int i = 0; i < span; i++) {
            Cell c = cells[i];
            switch (c.getT()) {
                case 's':
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

    //////////////////////////////////////Read Value///////////////////////////////////
    private String outOfBoundsMsg(int index) {
        return "Index: " + index + ", Size: " + span;
    }
    protected void rangeCheck(int index) {
        if (index >= span)
            throw new IndexOutOfBoundsException(outOfBoundsMsg(index));
    }

    protected Cell getCell(int i) {
        rangeCheck(i);
        return cells[i];
    }

    public boolean getBoolean(int columnIndex) {
        Cell c = getCell(columnIndex);
        boolean v;
        switch (c.getT()) {
            case 'b':
                v = c.getBv();
                break;
            case 'n':
            case 'd':
                v = c.getNv() != 0 || c.getDv() == 0.0;
                break;
            case 's':
                v = c.getSv() != null;
                break;
                default: v = false;
        }
        return v;
    }

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
            case 'd':
                b |= (int) c.getDv();
                break;
                default: throw new UncheckedTypeException("can't convert to byte");
        }
        return b;
    }

    public char getChar(int columnIndex) {
        Cell c = getCell(columnIndex);
        char cc = 0;
        switch (c.getT()) {
            case 's':
                String s = c.getSv();
                if (StringUtil.isNotEmpty(s)) {
                    cc |= s.charAt(0);
                }
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
                try {
                    l = Long.valueOf(c.getSv());
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

    public String getString(int columnIndex) {
        Cell c = getCell(columnIndex);
        String s;
        switch (c.getT()) {
            case 's':
            case 'l':
                s = c.getSv();
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

    public float getFloat(int columnIndex) {
        return (float) getDouble(columnIndex);
    }

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
            default: throw new UncheckedTypeException("unknown type");
        }
        return d;
    }

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
                date = DateUtil.toDate(c.getSv());
                break;
                default: throw new UncheckedTypeException("");
        }
        return date;
    }

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
                ts = DateUtil.toTimestamp(c.getSv());
                break;
            default: throw new UncheckedTypeException("");
        }
        return ts;
    }

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
        for (int i : columns) {
            if (fieldClazz[i] == String.class) {
                fields[i].set(t, getString(i));
            }
            else if (fieldClazz[i] == int.class || fieldClazz[i] == Integer.class) {
                fields[i].setInt(t, getInt(i));
            }
            else if (fieldClazz[i] == long.class || fieldClazz[i] == Long.class) {
                fields[i].setLong(t, getLong(i));
            }
            else if (fieldClazz[i] == java.util.Date.class || fieldClazz[i] == java.sql.Date.class) {
                fields[i].set(t, getDate(i));
            }
            else if (fieldClazz[i] == java.sql.Timestamp.class) {
                fields[i].set(t, getTimestamp(i));
            }
            else if (fieldClazz[i] == double.class || fieldClazz[i] == Double.class) {
                fields[i].setDouble(t, getDouble(i));
            }
            else if (fieldClazz[i] == float.class || fieldClazz[i] == Float.class) {
                fields[i].setFloat(t, getFloat(i));
            }
            else if (fieldClazz[i] == boolean.class || fieldClazz[i] == Boolean.class) {
                fields[i].setBoolean(t, getBoolean(i));
            }
            else if (fieldClazz[i] == char.class || fieldClazz[i] == Character.class) {
                fields[i].setChar(t, getChar(i));
            }
            else if (fieldClazz[i] == byte.class || fieldClazz[i] == Byte.class) {
                fields[i].setByte(t, getByte(i));
            }
            else if (fieldClazz[i] == short.class || fieldClazz[i] == Short.class) {
                fields[i].setShort(t, getShort(i));
            }
        }
    }

    ////////////////////////private inner class///////////////////////////////
    private static final class HeaderRow extends Row {
        String[] names;
        Class<?> clazz;
        Field[] fields;
        int[] columns;
        Class<?>[] fieldClazz;
        Object t;

        static final HeaderRow with(Row row) {
            HeaderRow hr = new HeaderRow();
            hr.names = new String[row.span];
            for (int i = 0; i < row.span; i++) {
                // header type is string
                hr.names[i] = row.cells[i].getSv();
            }
            return hr;
        }

        final boolean is(Class<?> clazz) {
            return this.clazz != null && this.clazz == clazz;
        }

        final HeaderRow setClass(Class<?> clazz) {
            this.clazz = clazz;
            Field[] fields = clazz.getDeclaredFields();
            int[] index = new int[fields.length];
            int count = 0;
            for (int i = 0, n = -1; i < fields.length; i++, n = -1) {
                Field f = fields[i];
                DisplayName ano = f.getAnnotation(DisplayName.class);
                if (ano != null) {
                    if (ano.skip()) {
                        fields[i] = null;
                        continue;
                    } else if (StringUtil.isNotEmpty(ano.value())) {
                        n = StringUtil.indexOf(names, ano.value());
                        if (n == -1) {
                            logger.warn(clazz + " field [" + ano.value() + "] can't find in header" + Arrays.toString(names));
                            fields[i] = null;
                            continue;
                        }
                    }
                }
                // no annotation or annotation value is null
                if (n < 0) {
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

        @Override
        public <T> T get(int columnIndex) {
            rangeCheck(columnIndex);
            return (T) names[columnIndex];
        }

        @Override public String toString() {
            StringJoiner joiner = new StringJoiner(" | ");
            for (String s : names) {
                joiner.add(s);
            }
            return joiner.toString();
        }
    }
}

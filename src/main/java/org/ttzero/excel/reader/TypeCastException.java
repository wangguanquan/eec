package org.ttzero.excel.reader;

import static org.ttzero.excel.entity.Sheet.int2Col;

/**
 * 读取时类型转换异常
 *
 * @author nasoda at 2024-02-23
 */
public class TypeCastException extends IllegalArgumentException {

    /**
     * 行号，从1开始
     */
    int row;

    /**
     * 列号，从1开始，可通过{@link #toColumnLetter()}转为字母
     */
    int col;

    CellType from;

    Class<?> to;

    public TypeCastException(int row, int col, CellType from, Class<?> to) {
        this.row = row;
        this.col = col;
        this.from = from;
        this.to = to;
    }

    public TypeCastException(int row, int col, CellType from, Class<?> to, String message) {
        super(message);
        this.row = row;
        this.col = col;
        this.from = from;
        this.to = to;
    }

    public TypeCastException(int row, int col, CellType from, Class<?> to, String message, Throwable cause) {
        super(message, cause);
        this.row = row;
        this.col = col;
        this.from = from;
        this.to = to;
    }

    public TypeCastException(int row, int col, CellType from, Class<?> to, Throwable cause) {
        super(cause);
        this.row = row;
        this.col = col;
        this.from = from;
        this.to = to;
    }

    public int getRow() {
        return row;
    }

    public void setRow(int row) {
        this.row = row;
    }

    public int getCol() {
        return col;
    }

    public void setCol(int col) {
        this.col = col;
    }

    public CellType getFrom() {
        return from;
    }

    public void setFrom(CellType from) {
        this.from = from;
    }

    public Class<?> getTo() {
        return to;
    }

    public void setTo(Class<?> to) {
        this.to = to;
    }

    public String toColumnLetter() {
        return new String(int2Col(col));
    }

}

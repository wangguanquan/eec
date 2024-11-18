package org.ttzero.excel.reader;

import static org.ttzero.excel.entity.Sheet.int2Col;
import static org.ttzero.excel.entity.Sheet.toCoordinate;

/**
 * 读取时类型转换异常
 *
 * @author nasoda on 2024-02-23
 */
public class TypeCastException extends IllegalArgumentException {

    /**
     * 行号，从1开始
     */
    public final int row;

    /**
     * 列号，从1开始，可通过{@link #toColumnLetter()}转为字母
     */
    public final int col;
    /**
     * Excel单元格类型
     */
    public final CellType from;
    /**
     * 目标转换类型
     */
    public final Class<?> to;

    public TypeCastException(int row, int col, CellType from, Class<?> to) {
        super("Can't cast " + from + "to " + to + " in cell '" + toCoordinate(row, col) + "'");
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

    public int getCol() {
        return col;
    }

    public CellType getFrom() {
        return from;
    }

    public Class<?> getTo() {
        return to;
    }

    public String toColumnLetter() {
        return new String(int2Col(col));
    }

}

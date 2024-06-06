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

import java.io.InputStream;
import java.math.BigDecimal;
import java.nio.ByteBuffer;
import java.nio.file.Path;

/**
 * 单元格，读取或写入时最小处理单元，它与上层的数据源无关
 *
 * @author guanquan.wang on 2018-09-22
 */
public class Cell {
    public Cell() { }
    public Cell(short i) { this.i = i; }
    public Cell(int i) { this.i = (short) (i & 0x7FFF); }
    public static final char SST          = 's';
    public static final char BOOL         = 'b';
    public static final char FUNCTION     = 'f';
    public static final char INLINESTR    = 'r';
    public static final char LONG         = 'l';
    public static final char DOUBLE       = 'd';
    public static final char NUMERIC      = 'n';
    public static final char BLANK        = 'k';
    public static final char CHARACTER    = 'c';
    public static final char DECIMAL      = 'm';
    public static final char DATETIME     = 'i';
    public static final char DATE         = 'a';
    public static final char TIME         = 't';
    public static final char UNALLOCATED  = '\0';
    public static final char EMPTY_TAG    = 'e';
    public static final char BINARY       = 'y';
    public static final char FILE         = 'x';
    public static final char INPUT_STREAM = 'p';
    public static final char REMOTE_URL   = 'u';
    public static final char BYTE_BUFFER  = 'o';
    /**
     * Unallocated cell
     */
    public static final Cell UNALLOCATED_CELL = new Cell();
    /**
     * Value type
     * n=numeric
     * s=string
     * b=boolean
     * f=function string
     * r=inlineStr
     * l=long
     * d=double
     */
    public char t; // type
    /**
     * String value
     */
    public String stringVal;
    /**
     * Integer value contain short
     */
    public int intVal;
    /**
     * Long value
     */
    public long longVal;
    /**
     * Double value contain float
     */
    public double doubleVal;
    /**
     * Boolean value
     */
    public boolean boolVal;
    /**
     * Char value
     */
    public char charVal;
    /**
     * Decimal value
     */
    public BigDecimal decimal;
    /**
     * Style index
     */
    public int xf;
    /**
     * Formula string
     */
    public String formula;
    /**
     * Shared calc id
     */
    public int si = -1;
    /**
     * Has formula
     */
    public boolean f;
    /**
     * Binary file (picture only)
     */
    public byte[] binary;
    /**
     * Binary file (picture only)
     */
    public ByteBuffer byteBuffer;
    /**
     * File path (picture file)
     */
    public Path path;
    /**
     * InputStream value (picture stream), auto-close after writen
     */
    public InputStream isv;
    /**
     * 是否为超链接
     */
    public boolean h;
    /**
     * 图片源类型
     */
    public char mediaType;
    /**
     * x-axis of cell in row
     */
    public transient short i;

    public Cell setT(char t) {
        this.t = t;
        return this;
    }

    public Cell setString(String sv) {
        this.t = INLINESTR;
        this.stringVal = sv;
        return this;
    }

    public Cell setInt(int nv) {
        this.t = NUMERIC;
        this.intVal = nv;
        return this;
    }

    public Cell setDouble(double dv) {
        this.t = DOUBLE;
        this.doubleVal = dv;
        return this;
    }

    public Cell setBool(boolean bv) {
        this.t = BOOL;
        this.boolVal = bv;
        return this;
    }

    public Cell setChar(char c) {
        this.t = CHARACTER;
        this.charVal = c;
        return this;
    }

    public Cell blank() {
        this.t = BLANK;
        return this;
    }

    public Cell emptyTag() {
        this.t = EMPTY_TAG;
        return this;
    }

    public Cell setLong(long lv) {
        this.t = LONG;
        this.longVal = lv;
        return this;
    }

    public Cell setDecimal(BigDecimal mv) {
        this.t = DECIMAL;
        this.decimal = mv;
        return this;
    }

    public Cell setDateTime(double i) {
        this.t = DATETIME;
        this.doubleVal = i;
        return this;
    }

    public Cell setDate(int a) {
        this.t = DATE;
        this.intVal = a;
        return this;
    }

    public Cell setTime(double t) {
        this.t = TIME;
        this.doubleVal = t;
        return this;
    }

    public Cell setBinary(byte[] bytes) {
        this.mediaType = BINARY;
        this.binary = bytes;
        return this;
    }

    public Cell setPath(Path path) {
        this.mediaType = FILE;
        this.path = path;
        return this;
    }

    public Cell setInputStream(InputStream stream) {
        this.mediaType = INPUT_STREAM;
        this.isv = stream;
        return this;
    }

    public Cell setByteBuffer(ByteBuffer byteBuffer) {
        this.mediaType = BYTE_BUFFER;
        this.byteBuffer = byteBuffer;
        return this;
    }

    public Cell setFormula(String formula) {
        if (formula != null && !formula.isEmpty()) {
            this.f = true;
            this.formula = formula;
        }
        return this;
    }

    public Cell setHyperlink(String hyperlink) {
        this.t = INLINESTR;
        this.stringVal = hyperlink;
        this.h = true;
        return this;
    }


    public Cell clear() {
        this.t  = UNALLOCATED;
        this.stringVal = null;
        this.intVal = 0;
        this.doubleVal = 0.0;
        this.boolVal = false;
        this.longVal = 0L;
        this.charVal = UNALLOCATED;
        this.decimal = null;
        this.xf = 0;
        this.formula = null;
        this.f = false;
        this.si = -1;
        this.binary = null;
        this.path = null;
        this.isv = null;
        this.byteBuffer = null;
        this.mediaType = UNALLOCATED;
        this.h = false;
        return this;
    }

    public Cell from(Cell cell) {
        this.t  = cell.t;
        this.stringVal = cell.stringVal;
        this.intVal = cell.intVal;
        this.doubleVal = cell.doubleVal;
        this.boolVal = cell.boolVal;
        this.longVal = cell.longVal;
        this.charVal = cell.charVal;
        this.decimal = cell.decimal;
        this.xf = cell.xf;
        this.formula = cell.formula;
        this.f = cell.f;
        this.si = cell.si;
        this.binary = cell.binary;
        this.path = cell.path;
        this.isv = cell.isv;
        this.byteBuffer = cell.byteBuffer;
        this.mediaType = cell.mediaType;
        this.h = cell.h;
        return this;
    }
}

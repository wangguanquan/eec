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

import java.math.BigDecimal;

/**
 * @author guanquan.wang on 2018-09-22
 */
public class Cell {
    public Cell() { }
    public Cell(short i) { this.i = i; }
    public Cell(int i) { this.i = (short) (i & 0x7FFF); }
    public static final char SST         = 's';
    public static final char BOOL        = 'b';
    public static final char FUNCTION    = 'f';
    public static final char INLINESTR   = 'r';
    public static final char LONG        = 'l';
    public static final char DOUBLE      = 'd';
    public static final char NUMERIC     = 'n';
    public static final char BLANK       = 'k';
    public static final char CHARACTER   = 'c';
    public static final char DECIMAL     = 'm';
    public static final char DATETIME    = 'i';
    public static final char DATE        = 'a';
    public static final char TIME        = 't';
    public static final char UNALLOCATED = '\0';
    public static final char EMPTY_TAG   = 'e';
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
    public String sv;
    /**
     * Integer value contain short
     */
    public int nv;
    /**
     * Long value
     */
    public long lv;
    /**
     * Double value contain float
     */
    public double dv;
    /**
     * Boolean value
     */
    public boolean bv;
    /**
     * Char value
     */
    public char cv;
    /**
     * Decimal value
     */
    public BigDecimal mv;
    /**
     * Style index
     */
    public int xf;
    /**
     * Formula string
     */
    public String fv;
    /**
     * Shared calc id
     */
    public int si;
    /**
     * Has formula
     */
    public boolean f;
    /**
     * x-axis of cell in row
     */
    public transient short i;

    public Cell setT(char t) {
        this.t = t;
        return this;
    }

    public Cell setSv(String sv) {
        this.t = INLINESTR;
        this.sv = sv;
        return this;
    }

    public Cell setNv(int nv) {
        this.t = NUMERIC;
        this.nv = nv;
        return this;
    }

    public Cell setDv(double dv) {
        this.t = DOUBLE;
        this.dv = dv;
        return this;
    }

    public Cell setBv(boolean bv) {
        this.t = BOOL;
        this.bv = bv;
        return this;
    }

    public Cell setCv(char c) {
        this.t = CHARACTER;
        this.cv = c;
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

    public Cell setLv(long lv) {
        this.t = LONG;
        this.lv = lv;
        return this;
    }

    public Cell setMv(BigDecimal mv) {
        this.t = DECIMAL;
        this.mv = mv;
        return this;
    }

    public Cell setIv(double i) {
        this.t = DATETIME;
        this.dv = i;
        return this;
    }

    public Cell setAv(int a) {
        this.t = DATE;
        this.nv = a;
        return this;
    }

    public Cell setTv(double t) {
        this.t = TIME;
        this.dv = t;
        return this;
    }

    public Cell clear() {
        this.t  = UNALLOCATED;
        this.sv = null;
        this.nv = 0;
        this.dv = 0.0;
        this.bv = false;
        this.lv = 0L;
        this.cv = '\0';
        this.mv = null;
        this.xf = 0;
        this.fv = null;
        this.f  = false;
        this.si = -1;
        return this;
    }

    public Cell from(Cell cell) {
        this.t  = cell.t;
        this.sv = cell.sv;
        this.nv = cell.nv;
        this.dv = cell.dv;
        this.bv = cell.bv;
        this.lv = cell.lv;
        this.cv = cell.cv;
        this.mv = cell.mv;
        this.xf = cell.xf;
        this.fv = cell.fv;
        this.f  = cell.f;
        this.si = cell.si;

        return this;
    }
}

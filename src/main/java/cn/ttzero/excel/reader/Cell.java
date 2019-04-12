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

/**
 * Create by guanquan.wang on 2018-09-22
 */
class Cell {
    Cell() {}
    private char t; // type
    // value
    private String sv;
    private int nv;
    private long lv;
    private double dv;
    private boolean bv;

    void setT(char t) {
        this.t = t;
    }

    void setSv(String sv) {
        this.sv = sv;
    }

    void setNv(int nv) {
        this.nv = nv;
    }

    void setDv(double dv) {
        this.dv = dv;
    }

    void setBv(boolean bv) {
        this.bv = bv;
    }

    long getLv() {
        return lv;
    }

    void setLv(long lv) {
        this.lv = lv;
    }

    char getT() {
        return t;
    }

    String getSv() {
        return sv;
    }

    int getNv() {
        return nv;
    }

    double getDv() {
        return dv;
    }

    boolean getBv() {
        return bv;
    }

    void clear() {
        this.t = '\0';
        this.sv = null;
        this.nv = 0;
        this.dv = 0.0;
        this.bv = false;
    }
}

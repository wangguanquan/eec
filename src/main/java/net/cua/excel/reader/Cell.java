package net.cua.excel.reader;

/**
 * Create by guanquan.wang at 2018-09-22
 */
class Cell {
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
        this.sv = null;
        this.nv = 0;
        this.dv = 0.0;
        this.bv = false;
    }
}

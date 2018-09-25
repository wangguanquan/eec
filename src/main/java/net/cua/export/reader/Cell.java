package net.cua.export.reader;

/**
 * Create by guanquan.wang at 2018-09-22
 */
public class Cell {
    private Class<?> clazz;
    private Object data;

    public Class<?> getClazz() {
        return clazz;
    }

    public void setClazz(Class<?> clazz) {
        this.clazz = clazz;
    }

    public Object getData() {
        return data;
    }

    public void setData(Object data) {
        this.data = data;
    }
}

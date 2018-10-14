package net.cua.excel.entity;

/**
 * Created by guanquan.wang at 2017/9/26.
 */
public class NameValue {
    private String name;
    private String value;

    public NameValue() {
    }

    public NameValue(String name, String value) {
        this.name = name;
        this.value = value;
    }

    public String getName() {
        return name;
    }

    public String getValue() {
        return value;
    }
}

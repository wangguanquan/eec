package net.cua.excel.entity.e3;

/**
 * Create by guanquan.wang at 2019-01-25 10:20
 */
public enum DocType {
    Empty(0x0),
    UserStorage(0x01),
    UserStream(0x02),
    LockBytes(0x03),
    Property(0x04),
    RootStorage(0x05)
    ;

    int value;

    DocType(int value) {
        this.value = value;
    }

    public int getValue() {
        return value;
    }

    public static DocType of(int value) {
        DocType[] types = values();
        if (value >= types.length || value < 0) return Empty;
        return types[value];
    }
}

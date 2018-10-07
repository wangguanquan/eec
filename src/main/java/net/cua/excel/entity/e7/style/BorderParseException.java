package net.cua.excel.entity.e7.style;

/**
 * Created by guanquan.wang at 2018-02-08 13:53
 */
public class BorderParseException extends RuntimeException {
    public BorderParseException(String message) {
        super(message);
    }

    public BorderParseException(String message, Throwable cause) {
        super(message, cause);
    }

    public BorderParseException(Throwable cause) {
        super(cause);
    }

    protected BorderParseException(String message, Throwable cause,
                                 boolean enableSuppression,
                                 boolean writableStackTrace) {
        super(message, cause, enableSuppression, writableStackTrace);
    }
}

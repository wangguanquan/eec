package net.cua.export.entity.e7.style;

/**
 * Created by guanquan.wang at 2018-02-05 16:10
 */
public class FontParseException extends RuntimeException {
    public FontParseException(String message) {
        super(message);
    }

    public FontParseException(String message, Throwable cause) {
        super(message, cause);
    }

    public FontParseException(Throwable cause) {
        super(cause);
    }

    protected FontParseException(String message, Throwable cause,
                        boolean enableSuppression,
                        boolean writableStackTrace) {
        super(message, cause, enableSuppression, writableStackTrace);
    }
}

package net.cua.export.entity.e7.style;

/**
 * Created by wanggq at 2018-02-09 14:25
 */
public class ColorParseException extends RuntimeException {
    public ColorParseException(String message) {
    super(message);
}

    public ColorParseException(String message, Throwable cause) {
        super(message, cause);
    }

    public ColorParseException(Throwable cause) {
        super(cause);
    }

    protected ColorParseException(String message, Throwable cause,
                                   boolean enableSuppression,
                                   boolean writableStackTrace) {
        super(message, cause, enableSuppression, writableStackTrace);
    }
}

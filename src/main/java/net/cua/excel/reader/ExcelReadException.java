package net.cua.excel.reader;

/**
 * Create by guanquan.wang at 2018-09-27 14:48
 */
public class ExcelReadException extends RuntimeException {
    public ExcelReadException(String message) {
        super(message);
    }

    public ExcelReadException(String message, Throwable cause) {
        super(message, cause);
    }

    public ExcelReadException(Throwable cause) {
        super(cause);
    }

    protected ExcelReadException(String message, Throwable cause,
                                 boolean enableSuppression,
                                 boolean writableStackTrace) {
        super(message, cause, enableSuppression, writableStackTrace);
    }
}

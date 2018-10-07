package net.cua.excel.reader;

import java.io.IOException;
import java.io.InvalidObjectException;
import java.io.ObjectInputStream;
import java.util.Objects;

/**
 * Create by guanquan.wang at 2018-09-30 11:59
 */
public class UncheckedTypeException extends RuntimeException {
    private static final long serialVersionUID = -8134305061645241065L;

    /**
     * Constructs an instance of this class.
     *
     * @param   message
     *          the detail message, can be null
     * @param   thr
     *          the {@code TypeException}
     *
     * @throws  NullPointerException
     *          if the cause is {@code null}
     */
    public UncheckedTypeException(String message, Throwable thr) {
        super(message, Objects.requireNonNull(thr));
    }

    /**
     * Constructs an instance of this class.
     *
     * @param   message the detail message
     *
     * @throws  NullPointerException
     *          if the cause is {@code null}
     */
    public UncheckedTypeException(String message) {
        super(message);
    }

    /**
     * Called to read the object from a stream.
     *
     * @throws InvalidObjectException
     *          if the object is invalid or has a cause that is not
     *          an {@code IOException}
     */
    private void readObject(ObjectInputStream s)
            throws IOException, ClassNotFoundException
    {
        s.defaultReadObject();
        Throwable cause = super.getCause();
        if (!(cause instanceof IOException))
            throw new InvalidObjectException("Cause must be an UncheckTypeException");
    }
}

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
     * @param message the detail message, can be null
     * @param thr     the {@code TypeException}
     * @throws NullPointerException if the cause is {@code null}
     */
    public UncheckedTypeException(String message, Throwable thr) {
        super(message, Objects.requireNonNull(thr));
    }

    /**
     * Constructs an instance of this class.
     *
     * @param message the detail message
     * @throws NullPointerException if the cause is {@code null}
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
        throws IOException, ClassNotFoundException {
        s.defaultReadObject();
        Throwable cause = super.getCause();
        if (!(cause instanceof IOException))
            throw new InvalidObjectException("Cause must be an UncheckTypeException");
    }
}

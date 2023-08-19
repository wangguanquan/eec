/*
 * Copyright (C) 2009 The Guava Authors
 *
 * Licensed under the Apache License, Version 2.0 (the "License"); you may not use this file except
 * in compliance with the License. You may obtain a copy of the License at
 *
 * http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software distributed under the License
 * is distributed on an "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express
 * or implied. See the License for the specific language governing permissions and limitations under
 * the License.
 */

package org.ttzero.excel.common.primitives;


/**
 * Static utility methods pertaining to {@code byte} primitives that interpret values as
 * <i>unsigned</i> (that is, any negative value {@code b} is treated as the positive value {@code
 * 256 + b}). The corresponding methods that treat the values as signed are found in {@code
 * SignedBytes}, and the methods for which signedness is not an issue are in {@code Bytes}.
 *
 * <p>See the Guava User Guide article on <a
 * href="https://github.com/google/guava/wiki/PrimitivesExplained">primitive utilities</a>.
 *
 * @author Kevin Bourrillion
 * @author Martin Buchholz
 * @author Hiroshi Yamauchi
 * @author Louis Wasserman
 * @since 1.0
 */
public final class UnsignedBytes {
  private UnsignedBytes() {}

  private static final int UNSIGNED_MASK = 0xFF;

  /**
   * Returns the value of the given byte as an integer, when treated as unsigned. That is, returns
   * {@code value + 256} if {@code value} is negative; {@code value} itself otherwise.
   *
   * <p><b>Java 8 users:</b> use {@link Byte#toUnsignedInt(byte)} instead.
   *
   * @since 6.0
   */
  public static int toInt(byte value) {
    return value & UNSIGNED_MASK;
  }

}

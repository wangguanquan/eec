/*
 * Copyright (C) 2011 The Guava Authors
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


import java.util.Arrays;
import java.util.Comparator;


/**
 * Static utility methods pertaining to {@code int} primitives that interpret values as
 * <i>unsigned</i> (that is, any negative value {@code x} is treated as the positive value {@code
 * 2^32 + x}). The methods for which signedness is not an issue are in {@link Ints}, as well as
 * signed versions of methods for which signedness is an issue.
 *
 * <p>In addition, this class provides several static methods for converting an {@code int} to a
 * {@code String} and a {@code String} to an {@code int} that treat the {@code int} as an unsigned
 * number.
 *
 * <p>Users of these utilities must be <i>extremely careful</i> not to mix up signed and unsigned
 * {@code int} values. When possible, it is recommended that the {@link UnsignedInteger} wrapper
 * class be used, at a small efficiency penalty, to enforce the distinction in the type system.
 *
 * <p>See the Guava User Guide article on <a
 * href="https://github.com/google/guava/wiki/PrimitivesExplained#unsigned-support">unsigned
 * primitive utilities</a>.
 *
 * @author Louis Wasserman
 * @since 11.0
 */
public final class UnsignedInts {
  static final long INT_MASK = 0xffffffffL;

  private UnsignedInts() {}

  static int flip(int value) {
    return value ^ Integer.MIN_VALUE;
  }

  /**
   * Compares the two specified {@code int} values, treating them as unsigned values between {@code
   * 0} and {@code 2^32 - 1} inclusive.
   *
   * <p><b>Java 8 users:</b> use {@link Integer#compareUnsigned(int, int)} instead.
   *
   * @param a the first unsigned {@code int} to compare
   * @param b the second unsigned {@code int} to compare
   * @return a negative value if {@code a} is less than {@code b}; a positive value if {@code a} is
   *     greater than {@code b}; or zero if they are equal
   */
  public static int compare(int a, int b) {
    return Integer.compare(flip(a), flip(b));
  }

  /**
   * Returns the value of the given {@code int} as a {@code long}, when treated as unsigned.
   *
   * <p><b>Java 8 users:</b> use {@link Integer#toUnsignedLong(int)} instead.
   */
  public static long toLong(int value) {
    return value & INT_MASK;
  }

}

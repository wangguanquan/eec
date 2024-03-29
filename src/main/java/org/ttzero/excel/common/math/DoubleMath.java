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

package org.ttzero.excel.common.math;

import static java.lang.Math.abs;
import static java.math.RoundingMode.HALF_EVEN;
import static java.math.RoundingMode.HALF_UP;

import java.math.RoundingMode;

/**
 * A class for arithmetic on doubles that is not covered by {@link Math}.
 *
 * @author Louis Wasserman
 * @since 11.0
 */
public final class DoubleMath {

  /**
   * Returns the result of dividing {@code p} by {@code q}, rounding using the specified {@code
   * RoundingMode}.
   *
   * @throws ArithmeticException if {@code q == 0}, or if {@code mode == UNNECESSARY} and {@code a}
   *     is not an integer multiple of {@code b}
   */
  @SuppressWarnings("fallthrough")
  public static long divide(long p, long q, RoundingMode mode) {
//    checkNotNull(mode);
    long div = p / q; // throws if q == 0
    long rem = p - q * div; // equals p % q

    if (rem == 0) {
      return div;
    }

    /*
     * Normal Java division rounds towards 0, consistently with RoundingMode.DOWN. We just have to
     * deal with the cases where rounding towards 0 is wrong, which typically depends on the sign of
     * p / q.
     *
     * signum is 1 if p and q are both nonnegative or both negative, and -1 otherwise.
     */
    int signum = 1 | (int) ((p ^ q) >> (Long.SIZE - 1));
    boolean increment;
    switch (mode) {
      case UNNECESSARY:
        // fall through
      case DOWN:
        increment = false;
        break;
      case UP:
        increment = true;
        break;
      case CEILING:
        increment = signum > 0;
        break;
      case FLOOR:
        increment = signum < 0;
        break;
      case HALF_EVEN:
      case HALF_DOWN:
      case HALF_UP:
        long absRem = abs(rem);
        long cmpRemToHalfDivisor = absRem - (abs(q) - absRem);
        // subtracting two nonnegative longs can't overflow
        // cmpRemToHalfDivisor has the same sign as compare(abs(rem), abs(q) / 2).
        if (cmpRemToHalfDivisor == 0) { // exactly on the half mark
          increment = (mode == HALF_UP | (mode == HALF_EVEN & (div & 1) != 0));
        } else {
          increment = cmpRemToHalfDivisor > 0; // closer to the UP value
        }
        break;
      default:
        throw new AssertionError();
    }
    return increment ? div + signum : div;
  }


  private DoubleMath() {}
}

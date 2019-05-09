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

package com.google.common.math;

import static com.google.common.math.DoubleUtils.SIGNIFICAND_BITS;
import static com.google.common.math.DoubleUtils.getSignificand;
import static com.google.common.math.DoubleUtils.isFinite;
import static java.lang.Math.abs;
import static java.lang.Math.copySign;
import static java.lang.Math.getExponent;
import static java.lang.Math.rint;
import static java.math.RoundingMode.HALF_EVEN;
import static java.math.RoundingMode.HALF_UP;

import java.math.RoundingMode;
import java.util.Iterator;

/**
 * A class for arithmetic on doubles that is not covered by {@link Math}.
 *
 * @author Louis Wasserman
 * @since 11.0
 */
public final class DoubleMath {
  /*
   * This method returns a value y such that rounding y DOWN (towards zero) gives the same result as
   * rounding x according to the specified mode.
   */
  static double roundIntermediate(double x, RoundingMode mode) {
    if (!isFinite(x)) {
      throw new ArithmeticException("input is infinite or NaN");
    }
    switch (mode) {
      case UNNECESSARY:
        return x;

      case FLOOR:
        if (x >= 0.0 || isMathematicalInteger(x)) {
          return x;
        } else {
          return (long) x - 1;
        }

      case CEILING:
        if (x <= 0.0 || isMathematicalInteger(x)) {
          return x;
        } else {
          return (long) x + 1;
        }

      case DOWN:
        return x;

      case UP:
        if (isMathematicalInteger(x)) {
          return x;
        } else {
          return (long) x + (x > 0 ? 1 : -1);
        }

      case HALF_EVEN:
        return rint(x);

      case HALF_UP:
        {
          double z = rint(x);
          if (abs(x - z) == 0.5) {
            return x + copySign(0.5, x);
          } else {
            return z;
          }
        }

      case HALF_DOWN:
        {
          double z = rint(x);
          if (abs(x - z) == 0.5) {
            return x;
          } else {
            return z;
          }
        }

      default:
        throw new AssertionError();
    }
  }


  /**
   * Returns the {@code long} value that is equal to {@code x} rounded with the specified rounding
   * mode, if possible.
   *
   * @throws ArithmeticException if
   *     <ul>
   *       <li>{@code x} is infinite or NaN
   *       <li>{@code x}, after being rounded to a mathematical integer using the specified rounding
   *           mode, is either less than {@code Long.MIN_VALUE} or greater than {@code
   *           Long.MAX_VALUE}
   *       <li>{@code x} is not a mathematical integer and {@code mode} is {@link
   *           RoundingMode#UNNECESSARY}
   *     </ul>
   */
  public static long roundToLong(double x, RoundingMode mode) {
    double z = roundIntermediate(x, mode);
//    checkInRangeForRoundingInputs(
//        MIN_LONG_AS_DOUBLE - z < 1.0 & z < MAX_LONG_AS_DOUBLE_PLUS_ONE, x, mode);
    return (long) z;
  }


  /**
   * Returns {@code true} if {@code x} is exactly equal to {@code 2^k} for some finite integer
   * {@code k}.
   */
  public static boolean isPowerOfTwo(double x) {
    if (x > 0.0 && isFinite(x)) {
      long significand = getSignificand(x);
      return (significand & (significand - 1)) == 0;
    }
    return false;
  }

  /**
   * Returns {@code true} if {@code x} represents a mathematical integer.
   *
   * <p>This is equivalent to, but not necessarily implemented as, the expression {@code
   * !Double.isNaN(x) && !Double.isInfinite(x) && x == Math.rint(x)}.
   */
  public static boolean isMathematicalInteger(double x) {
    return isFinite(x)
        && (x == 0.0
            || SIGNIFICAND_BITS - Long.numberOfTrailingZeros(getSignificand(x)) <= getExponent(x));
  }


  /**
   * Returns the <a href="http://en.wikipedia.org/wiki/Arithmetic_mean">arithmetic mean</a> of
   * {@code values}.
   *
   * <p>If these values are a sample drawn from a population, this is also an unbiased estimator of
   * the arithmetic mean of the population.
   *
   * @param values a nonempty series of values
   * @throws IllegalArgumentException if {@code values} is empty or contains any non-finite value
   * @deprecated Use instead, noting the less strict handling of non-finite
   *     values.
   */
  // com.google.common.math.DoubleUtils
  public static double mean(double... values) {
//    checkArgument(values.length > 0, "Cannot take mean of 0 values");
    long count = 1;
    double mean = checkFinite(values[0]);
    for (int index = 1; index < values.length; ++index) {
      checkFinite(values[index]);
      count++;
      // Art of Computer Programming vol. 2, Knuth, 4.2.2, (15)
      mean += (values[index] - mean) / count;
    }
    return mean;
  }

  /**
   * Returns the <a href="http://en.wikipedia.org/wiki/Arithmetic_mean">arithmetic mean</a> of
   * {@code values}.
   *
   * <p>If these values are a sample drawn from a population, this is also an unbiased estimator of
   * the arithmetic mean of the population.
   *
   * @param values a nonempty series of values
   * @throws IllegalArgumentException if {@code values} is empty
   * @deprecated Use instead, noting the less strict handling of non-finite
   *     values.
   */
  @Deprecated
  public static double mean(int... values) {
//    checkArgument(values.length > 0, "Cannot take mean of 0 values");
    // The upper bound on the the length of an array and the bounds on the int values mean that, in
    // this case only, we can compute the sum as a long without risking overflow or loss of
    // precision. So we do that, as it's slightly quicker than the Knuth algorithm.
    long sum = 0;
    for (int index = 0; index < values.length; ++index) {
      sum += values[index];
    }
    return (double) sum / values.length;
  }

  /**
   * Returns the <a href="http://en.wikipedia.org/wiki/Arithmetic_mean">arithmetic mean</a> of
   * {@code values}.
   *
   * <p>If these values are a sample drawn from a population, this is also an unbiased estimator of
   * the arithmetic mean of the population.
   *
   * @param values a nonempty series of values, which will be converted to {@code double} values
   *     (this may cause loss of precision for longs of magnitude over 2^53 (slightly over 9e15))
   * @throws IllegalArgumentException if {@code values} is empty
   * @deprecated Use  instead, noting the less strict handling of non-finite
   *     values.
   */
  @Deprecated
  public static double mean(long... values) {
//    checkArgument(values.length > 0, "Cannot take mean of 0 values");
    long count = 1;
    double mean = values[0];
    for (int index = 1; index < values.length; ++index) {
      count++;
      // Art of Computer Programming vol. 2, Knuth, 4.2.2, (15)
      mean += (values[index] - mean) / count;
    }
    return mean;
  }

  /**
   * Returns the <a href="http://en.wikipedia.org/wiki/Arithmetic_mean">arithmetic mean</a> of
   * {@code values}.
   *
   * <p>If these values are a sample drawn from a population, this is also an unbiased estimator of
   * the arithmetic mean of the population.
   *
   * @param values a nonempty series of values, which will be converted to {@code double} values
   *     (this may cause loss of precision)
   * @throws IllegalArgumentException if {@code values} is empty or contains any non-finite value
   * @deprecated Use instead, noting the less strict handling of non-finite
   *     values.
   */
  @Deprecated
  // com.google.common.math.DoubleUtils
  public static double mean(Iterable<? extends Number> values) {
    return mean(values.iterator());
  }

  /**
   * Returns the <a href="http://en.wikipedia.org/wiki/Arithmetic_mean">arithmetic mean</a> of
   * {@code values}.
   *
   * <p>If these values are a sample drawn from a population, this is also an unbiased estimator of
   * the arithmetic mean of the population.
   *
   * @param values a nonempty series of values, which will be converted to {@code double} values
   *     (this may cause loss of precision)
   * @throws IllegalArgumentException if {@code values} is empty or contains any non-finite value
   * @deprecated Use instead, noting the less strict handling of non-finite
   *     values.
   */
  @Deprecated
  // com.google.common.math.DoubleUtils
  public static double mean(Iterator<? extends Number> values) {
//    checkArgument(values.hasNext(), "Cannot take mean of 0 values");
    long count = 1;
    double mean = checkFinite(values.next().doubleValue());
    while (values.hasNext()) {
      double value = checkFinite(values.next().doubleValue());
      count++;
      // Art of Computer Programming vol. 2, Knuth, 4.2.2, (15)
      mean += (value - mean) / count;
    }
    return mean;
  }

  private static double checkFinite(double argument) {
//    checkArgument(isFinite(argument));
    return argument;
  }

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

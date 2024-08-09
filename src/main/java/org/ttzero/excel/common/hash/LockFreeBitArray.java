package org.ttzero.excel.common.hash;

import java.math.RoundingMode;
import java.util.Arrays;
import java.util.concurrent.atomic.AtomicLongArray;
import java.util.concurrent.atomic.LongAdder;

import static java.lang.Math.abs;
import static java.math.RoundingMode.HALF_EVEN;
import static java.math.RoundingMode.HALF_UP;

/**
 * Models a lock-free array of bits.
 *
 * <p>We use this instead of java.util.BitSet because we need access to the array of longs and we
 * need compare-and-swap.
 */
public final class LockFreeBitArray {
    private static final int LONG_ADDRESSABLE_BITS = 6;
    final AtomicLongArray data;
    private final LongAdder bitCount;

    LockFreeBitArray(long bits) {
        this(new long[(int) (divide(bits, 64, RoundingMode.CEILING))]);
    }

    // Used by serialization
    LockFreeBitArray(long[] data) {
        this.data = new AtomicLongArray(data);
        this.bitCount = new LongAdder();
        long bitCount = 0;
        for (long value : data) {
            bitCount += Long.bitCount(value);
        }
        this.bitCount.add(bitCount);
    }

    /** Returns true if the bit changed value. */
    boolean set(long bitIndex) {
        if (get(bitIndex)) {
            return false;
        }

        int longIndex = (int) (bitIndex >>> LONG_ADDRESSABLE_BITS);
        long mask = 1L << bitIndex; // only cares about low 6 bits of bitIndex

        long oldValue;
        long newValue;
        do {
            oldValue = data.get(longIndex);
            newValue = oldValue | mask;
            if (oldValue == newValue) {
                return false;
            }
        } while (!data.compareAndSet(longIndex, oldValue, newValue));

        // We turned the bit on, so increment bitCount.
        bitCount.increment();
        return true;
    }

    boolean get(long bitIndex) {
        return (data.get((int) (bitIndex >>> LONG_ADDRESSABLE_BITS)) & (1L << bitIndex)) != 0;
    }

    /**
     * Careful here: if threads are mutating the atomicLongArray while this method is executing, the
     * final long[] will be a "rolling snapshot" of the state of the bit array. This is usually good
     * enough, but should be kept in mind.
     */
    public static long[] toPlainArray(AtomicLongArray atomicLongArray) {
        long[] array = new long[atomicLongArray.length()];
        for (int i = 0; i < array.length; ++i) {
            array[i] = atomicLongArray.get(i);
        }
        return array;
    }

    /** Number of bits */
    long bitSize() {
        return (long) data.length() * Long.SIZE;
    }

    LockFreeBitArray copy() {
        return new LockFreeBitArray(toPlainArray(data));
    }

    /**
     * Combines the two BitArrays using bitwise OR.
     *
     * <p>NOTE: Because of the use of atomics, if the other LockFreeBitArray is being mutated while
     * this operation is executing, not all of those new 1's may be set in the final state of this
     * LockFreeBitArray. The ONLY guarantee provided is that all the bits that were set in the other
     * LockFreeBitArray at the start of this method will be set in this LockFreeBitArray at the end
     * of this method.
     */
    void putAll(LockFreeBitArray other) {
        for (int i = 0; i < data.length(); i++) {
            long otherLong = other.data.get(i);

            long ourLongOld;
            long ourLongNew;
            boolean changedAnyBits = true;
            do {
                ourLongOld = data.get(i);
                ourLongNew = ourLongOld | otherLong;
                if (ourLongOld == ourLongNew) {
                    changedAnyBits = false;
                    break;
                }
            } while (!data.compareAndSet(i, ourLongOld, ourLongNew));

            if (changedAnyBits) {
                int bitsAdded = Long.bitCount(ourLongNew) - Long.bitCount(ourLongOld);
                bitCount.add(bitsAdded);
            }
        }
    }

    @Override
    public boolean equals(Object o) {
        if (o instanceof LockFreeBitArray) {
            LockFreeBitArray lockFreeBitArray = (LockFreeBitArray) o;
            return Arrays.equals(toPlainArray(data), toPlainArray(lockFreeBitArray.data));
        }
        return false;
    }

    @Override
    public int hashCode() {
        return Arrays.hashCode(toPlainArray(data));
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
}
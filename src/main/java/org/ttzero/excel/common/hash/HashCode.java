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

package org.ttzero.excel.common.hash;


import java.io.Serializable;


/**
 * An immutable hash code of arbitrary bit length.
 *
 * @author Dimitris Andreou
 * @author Kurt Alfred Kluever
 * @since 11.0
 */
public abstract class HashCode {
  HashCode() {}

  /** Returns the number of bits in this hash code; a positive multiple of 8. */
  public abstract int bits();

  /**
   * Returns the first four bytes of {@linkplain #asBytes() this hashcode's bytes}, converted to an
   * {@code int} value in little-endian order.
   *
   * @throws IllegalStateException if {@code bits() < 32}
   */
  public abstract int asInt();

  /**
   * Returns the value of this hash code as a byte array. The caller may modify the byte array;
   * changes to it will <i>not</i> be reflected in this {@code HashCode} object or any other arrays
   * returned by this method.
   */
  // TODO(user): consider ByteString here, when that is available
  public abstract byte[] asBytes();

  /**
   * Returns a mutable view of the underlying bytes for the given {@code HashCode} if it is a
   * byte-based hashcode. Otherwise it returns {@link HashCode#asBytes}. Do <i>not</i> mutate this
   * array or else you will break the immutability contract of {@code HashCode}.
   */
  byte[] getBytesInternal() {
    return asBytes();
  }

  /**
   * Returns whether this {@code HashCode} and that {@code HashCode} have the same value, given that
   * they have the same number of bits.
   */
  abstract boolean equalsSameBits(HashCode that);

  /**
   * Creates a 32-bit {@code HashCode} representation of the given int value. The underlying bytes
   * are interpreted in little endian order.
   *
   * @since 15.0 (since 12.0 in HashCodes)
   */
  public static HashCode fromInt(int hash) {
    return new IntHashCode(hash);
  }

  private static final class IntHashCode extends HashCode implements Serializable {
    final int hash;

    IntHashCode(int hash) {
      this.hash = hash;
    }

    @Override
    public int bits() {
      return 32;
    }

    @Override
    public byte[] asBytes() {
      return new byte[] {(byte) hash, (byte) (hash >> 8), (byte) (hash >> 16), (byte) (hash >> 24)};
    }

    @Override
    public int asInt() {
      return hash;
    }

    @Override
    boolean equalsSameBits(HashCode that) {
      return hash == that.asInt();
    }

    private static final long serialVersionUID = 0;
  }

  /**
   * Creates a {@code HashCode} from a byte array. The array is <i>not</i> copied defensively, so it
   * must be handed-off so as to preserve the immutability contract of {@code HashCode}.
   */
  static HashCode fromBytesNoCopy(byte[] bytes) {
    return new BytesHashCode(bytes);
  }

  private static final class BytesHashCode extends HashCode implements Serializable {
    final byte[] bytes;

    BytesHashCode(byte[] bytes) {
      this.bytes = bytes;
    }

    @Override
    public int bits() {
      return bytes.length * 8;
    }

    @Override
    public byte[] asBytes() {
      return bytes.clone();
    }

    @Override
    public int asInt() {
      return (bytes[0] & 0xFF)
          | ((bytes[1] & 0xFF) << 8)
          | ((bytes[2] & 0xFF) << 16)
          | ((bytes[3] & 0xFF) << 24);
    }

    @Override
    byte[] getBytesInternal() {
      return bytes;
    }

    @Override
    boolean equalsSameBits(HashCode that) {
      // We don't use MessageDigest.isEqual() here because its contract does not guarantee
      // constant-time evaluation (no short-circuiting).
      if (this.bytes.length != that.getBytesInternal().length) {
        return false;
      }

      boolean areEqual = true;
      for (int i = 0; i < this.bytes.length; i++) {
        areEqual &= (this.bytes[i] == that.getBytesInternal()[i]);
      }
      return areEqual;
    }

    private static final long serialVersionUID = 0;
  }

  /**
   * Returns {@code true} if {@code object} is a {@link HashCode} instance with the identical byte
   * representation to this hash code.
   *
   * <p><b>Security note:</b> this method uses a constant-time (not short-circuiting) implementation
   * to protect against <a href="http://en.wikipedia.org/wiki/Timing_attack">timing attacks</a>.
   */
  @Override
  public final boolean equals(Object object) {
    if (object instanceof HashCode) {
      HashCode that = (HashCode) object;
      return bits() == that.bits() && equalsSameBits(that);
    }
    return false;
  }

  /**
   * Returns a "Java hash code" for this {@code HashCode} instance; this is well-defined (so, for
   * example, you can safely put {@code HashCode} instances into a {@code HashSet}) but is otherwise
   * probably not what you want to use.
   */
  @Override
  public final int hashCode() {
    // If we have at least 4 bytes (32 bits), just take the first 4 bytes. Since this is
    // already a (presumably) high-quality hash code, any four bytes of it will do.
    if (bits() >= 32) {
      return asInt();
    }
    // If we have less than 4 bytes, use them all.
    byte[] bytes = getBytesInternal();
    int val = (bytes[0] & 0xFF);
    for (int i = 1; i < bytes.length; i++) {
      val |= ((bytes[i] & 0xFF) << (i * 8));
    }
    return val;
  }

  /**
   * Returns a string containing each byte of {@link #asBytes}, in order, as a two-digit unsigned
   * hexadecimal number in lower case.
   *
   * <p>Note that if the output is considered to be a single hexadecimal number, this hash code's
   * bytes are the <i>big-endian</i> representation of that number. This may be surprising since
   * everything else in the hashing API uniformly treats multibyte values as little-endian. But this
   * format conveniently matches that of utilities such as the UNIX {@code md5sum} command.
   *
   * <p>To create a {@code HashCode} from its string representation, see {@code #fromString}.
   */
  @Override
  public final String toString() {
    byte[] bytes = getBytesInternal();
    StringBuilder sb = new StringBuilder(2 * bytes.length);
    for (byte b : bytes) {
      sb.append(hexDigits[(b >> 4) & 0xf]).append(hexDigits[b & 0xf]);
    }
    return sb.toString();
  }

  private static final char[] hexDigits = "0123456789abcdef".toCharArray();
}

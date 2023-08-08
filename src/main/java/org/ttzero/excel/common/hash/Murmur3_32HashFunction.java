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

/*
 * MurmurHash3 was written by Austin Appleby, and is placed in the public
 * domain. The author hereby disclaims copyright to this source code.
 */

/*
 * Source:
 * http://code.google.com/p/smhasher/source/browse/trunk/MurmurHash3.cpp
 * (Modified to adapt to Guava coding conventions and to use the HashFunction interface)
 */

package org.ttzero.excel.common.hash;

import java.io.Serializable;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;

/**
 * See MurmurHash3_x86_32 in <a
 * href="https://github.com/aappleby/smhasher/blob/master/src/MurmurHash3.cpp">the C++
 * implementation</a>.
 *
 * @author Austin Appleby
 * @author Dimitris Andreou
 * @author Kurt Alfred Kluever
 */
final class Murmur3_32HashFunction extends AbstractHashFunction implements Serializable {
  static final HashFunction MURMUR3_32 = new Murmur3_32HashFunction(0);

  private static final int C1 = 0xcc9e2d51;
  private static final int C2 = 0x1b873593;

  private final int seed;

  Murmur3_32HashFunction(int seed) {
    this.seed = seed;
  }

  @Override
  public int bits() {
    return 32;
  }

  @Override
  public Hasher newHasher() {
    return new Murmur3_32Hasher(seed);
  }

  @Override
  public String toString() {
    return "Hashing.murmur3_32(" + seed + ")";
  }

  @Override
  public boolean equals( Object object) {
    if (object instanceof Murmur3_32HashFunction) {
      Murmur3_32HashFunction other = (Murmur3_32HashFunction) object;
      return seed == other.seed;
    }
    return false;
  }

  @Override
  public int hashCode() {
    return getClass().hashCode() ^ seed;
  }

  private static int getIntLittleEndian(byte[] input, int offset) {
    return fromBytes(input[offset + 3], input[offset + 2], input[offset + 1], input[offset]);
  }

  /**
   * Returns the {@code int} value whose byte representation is the given 4 bytes, in big-endian
   * order; equivalent to {@code Ints.fromByteArray(new byte[] {b1, b2, b3, b4})}.
   *
   * @since 7.0
   */
  public static int fromBytes(byte b1, byte b2, byte b3, byte b4) {
    return b1 << 24 | (b2 & 0xFF) << 16 | (b3 & 0xFF) << 8 | (b4 & 0xFF);
  }

  private static int mixK1(int k1) {
    k1 *= C1;
    k1 = Integer.rotateLeft(k1, 15);
    k1 *= C2;
    return k1;
  }

  private static int mixH1(int h1, int k1) {
    h1 ^= k1;
    h1 = Integer.rotateLeft(h1, 13);
    h1 = h1 * 5 + 0xe6546b64;
    return h1;
  }

  // Finalization mix - force all bits of a hash block to avalanche
  private static HashCode fmix(int h1, int length) {
    h1 ^= length;
    h1 ^= h1 >>> 16;
    h1 *= 0x85ebca6b;
    h1 ^= h1 >>> 13;
    h1 *= 0xc2b2ae35;
    h1 ^= h1 >>> 16;
    return HashCode.fromInt(h1);
  }

  private static final class Murmur3_32Hasher extends AbstractHasher {
    private int h1;
    private long buffer;
    private int shift;
    private int length;

    Murmur3_32Hasher(int seed) {
      this.h1 = seed;
      this.length = 0;
    }

    private void update(int nBytes, long update) {
      // 1 <= nBytes <= 4
      buffer |= (update & 0xFFFFFFFFL) << shift;
      shift += nBytes * 8;
      length += nBytes;

      if (shift >= 32) {
        h1 = mixH1(h1, mixK1((int) buffer));
        buffer >>>= 32;
        shift -= 32;
      }
    }

    @Override
    public Hasher putByte(byte b) {
      update(1, b & 0xFF);
      return this;
    }

    @Override
    public Hasher putBytes(byte[] bytes, int off, int len) {
//      checkPositionIndexes(off, off + len, bytes.length);
      int i;
      for (i = 0; i + 4 <= len; i += 4) {
        update(4, getIntLittleEndian(bytes, off + i));
      }
      for (; i < len; i++) {
        putByte(bytes[off + i]);
      }
      return this;
    }

    @SuppressWarnings("deprecation") // need to use Charsets for Android tests to pass
    @Override
    public Hasher putString(CharSequence input, Charset charset) {
      if (StandardCharsets.UTF_8.equals(charset)) {
        int utf16Length = input.length();
        int i = 0;

        // This loop optimizes for pure ASCII.
        while (i + 4 <= utf16Length) {
          char c0 = input.charAt(i);
          char c1 = input.charAt(i + 1);
          char c2 = input.charAt(i + 2);
          char c3 = input.charAt(i + 3);
          if (c0 < 0x80 && c1 < 0x80 && c2 < 0x80 && c3 < 0x80) {
            update(4, c0 | (c1 << 8) | (c2 << 16) | (c3 << 24));
            i += 4;
          } else {
            break;
          }
        }

        for (; i < utf16Length; i++) {
          char c = input.charAt(i);
          if (c < 0x80) {
            update(1, c);
          } else if (c < 0x800) {
            update(2, charToTwoUtf8Bytes(c));
          } else if (c < Character.MIN_SURROGATE || c > Character.MAX_SURROGATE) {
            update(3, charToThreeUtf8Bytes(c));
          } else {
            int codePoint = Character.codePointAt(input, i);
            if (codePoint == c) {
              // fall back to JDK getBytes instead of trying to handle invalid surrogates ourselves
              putBytes(input.subSequence(i, utf16Length).toString().getBytes(charset));
              return this;
            }
            i++;
            update(4, codePointToFourUtf8Bytes(codePoint));
          }
        }
        return this;
      } else {
        return super.putString(input, charset);
      }
    }

    @Override
    public HashCode hash() {
//      checkState(!isDone);
      h1 ^= mixK1((int) buffer);
      return fmix(h1, length);
    }
  }

  private static long codePointToFourUtf8Bytes(int codePoint) {
    return (((0xFL << 4) | (codePoint >>> 18)) & 0xFF)
        | ((0x80L | (0x3F & (codePoint >>> 12))) << 8)
        | ((0x80L | (0x3F & (codePoint >>> 6))) << 16)
        | ((0x80L | (0x3F & codePoint)) << 24);
  }

  private static long charToThreeUtf8Bytes(char c) {
    return (((0xF << 5) | (c >>> 12)) & 0xFF)
        | ((0x80 | (0x3F & (c >>> 6))) << 8)
        | ((0x80 | (0x3F & c)) << 16);
  }

  private static long charToTwoUtf8Bytes(char c) {
    return (((0xF << 6) | (c >>> 6)) & 0xFF) | ((0x80 | (0x3F & c)) << 8);
  }

  private static final long serialVersionUID = 0L;
}

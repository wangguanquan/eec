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


/**
 * Static methods to obtain {@link HashFunction} instances, and other static hashing-related
 * utilities.
 *
 * <p>A comparison of the various hash functions can be found <a
 * href="http://goo.gl/jS7HH">here</a>.
 *
 * @author Kevin Bourrillion
 * @author Dimitris Andreou
 * @author Kurt Alfred Kluever
 * @since 11.0
 */
public final class Hashing {

  /**
   * Returns a hash function implementing the <a
   * href="https://github.com/aappleby/smhasher/blob/master/src/MurmurHash3.cpp">32-bit murmur3
   * algorithm, x86 variant</a> (little-endian variant), using a seed value of zero.
   *
   * <p>The exact C++ equivalent is the MurmurHash3_x86_32 function (Murmur3A).
   */
  public static HashFunction murmur3_32() {
    return Murmur3_32HashFunction.MURMUR3_32;
  }

  /**
   * Returns a hash function implementing the <a
   * href="https://github.com/aappleby/smhasher/blob/master/src/MurmurHash3.cpp">128-bit murmur3
   * algorithm, x64 variant</a> (little-endian variant), using the given seed value.
   *
   * <p>The exact C++ equivalent is the MurmurHash3_x64_128 function (Murmur3F).
   */
  public static HashFunction murmur3_128(int seed) {
    return new Murmur3_128HashFunction(seed);
  }

  /**
   * Returns a hash function implementing the <a
   * href="https://github.com/aappleby/smhasher/blob/master/src/MurmurHash3.cpp">128-bit murmur3
   * algorithm, x64 variant</a> (little-endian variant), using a seed value of zero.
   *
   * <p>The exact C++ equivalent is the MurmurHash3_x64_128 function (Murmur3F).
   */
  public static HashFunction murmur3_128() {
    return Murmur3_128HashFunction.MURMUR3_128;
  }

  /**
   * Returns a hash function implementing the MD5 hash algorithm (128 hash bits).
   *
   * @deprecated If you must interoperate with a system that requires MD5, then use this method,
   *     despite its deprecation. But if you can choose your hash function, avoid MD5, which is
   *     neither fast nor secure. As of January 2017, we suggest:
   *     <ul>
   *       <li>For security:
   *           {@code Hashing#sha256} or a higher-level API.
   *       <li>For speed: {@code Hashing#goodFastHash}, though see its docs for caveats.
   *     </ul>
   */
  @Deprecated
  public static HashFunction md5() {
    return Md5Holder.MD5;
  }

  private static class Md5Holder {
    static final HashFunction MD5 = new MessageDigestHashFunction("MD5", "Hashing.md5()");
  }

  /**
   * Returns a hash function implementing the SHA-1 algorithm (160 hash bits).
   *
   * @deprecated If you must interoperate with a system that requires SHA-1, then use this method,
   *     despite its deprecation. But if you can choose your hash function, avoid SHA-1, which is
   *     neither fast nor secure. As of January 2017, we suggest:
   *     <ul>
   *       <li>For security:
   *           {@code Hashing#sha256} or a higher-level API.
   *       <li>For speed: {@code Hashing#goodFastHash}, though see its docs for caveats.
   *     </ul>
   */
  @Deprecated
  public static HashFunction sha1() {
    return Sha1Holder.SHA_1;
  }

  private static class Sha1Holder {
    static final HashFunction SHA_1 = new MessageDigestHashFunction("SHA-1", "Hashing.sha1()");
  }

  /** Returns a hash function implementing the SHA-512 algorithm (512 hash bits). */
  public static HashFunction sha512() {
    return Sha512Holder.SHA_512;
  }

  private static class Sha512Holder {
    static final HashFunction SHA_512 =
        new MessageDigestHashFunction("SHA-512", "Hashing.sha512()");
  }

  private Hashing() {}
}

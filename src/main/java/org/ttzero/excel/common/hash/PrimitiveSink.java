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


import java.nio.charset.Charset;

/**
 * An object which can receive a stream of primitive values.
 *
 * @author Kevin Bourrillion
 * @since 12.0 (in 11.0 as {@code Sink})
 */
public interface PrimitiveSink {
  /**
   * Puts a byte into this sink.
   *
   * @param b a byte
   * @return this instance
   */
  PrimitiveSink putByte(byte b);

  /**
   * Puts an array of bytes into this sink.
   *
   * @param bytes a byte array
   * @return this instance
   */
  PrimitiveSink putBytes(byte[] bytes);

  /**
   * Puts a chunk of an array of bytes into this sink. {@code bytes[off]} is the first byte written,
   * {@code bytes[off + len - 1]} is the last.
   *
   * @param bytes a byte array
   * @param off the start offset in the array
   * @param len the number of bytes to write
   * @return this instance
   * @throws IndexOutOfBoundsException if {@code off < 0} or {@code off + len > bytes.length} or
   *     {@code len < 0}
   */
  PrimitiveSink putBytes(byte[] bytes, int off, int len);

  /**
   * Puts a string into this sink using the given charset.
   *
   * <p><b>Warning:</b> This method, which reencodes the input before processing it, is useful only
   * for cross-language compatibility. For other use cases, prefer {@code #putUnencodedChars}, which
   * is faster, produces the same output across Java releases, and processes every {@code char} in
   * the input, even if some are invalid.
   */
  PrimitiveSink putString(CharSequence charSequence, Charset charset);
}

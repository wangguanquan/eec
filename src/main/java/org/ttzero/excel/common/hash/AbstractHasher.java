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
 * An abstract implementation of {@link Hasher}, which only requires subtypes to implement {@link
 * #putByte}. Subtypes may provide more efficient implementations, however.
 *
 * @author Dimitris Andreou
 */
abstract class AbstractHasher implements Hasher {

  @Override
  public Hasher putString(CharSequence charSequence, Charset charset) {
    return putBytes(charSequence.toString().getBytes(charset));
  }

  @Override
  public Hasher putBytes(byte[] bytes) {
    return putBytes(bytes, 0, bytes.length);
  }

  @Override
  public Hasher putBytes(byte[] bytes, int off, int len) {
    for (int i = 0; i < len; i++) {
      putByte(bytes[off + i]);
    }
    return this;
  }

  @Override
  public <T> Hasher putObject(T instance, Funnel<? super T> funnel) {
    funnel.funnel(instance, this);
    return this;
  }
}

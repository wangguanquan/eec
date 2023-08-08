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
import java.nio.charset.Charset;

/**
 * Funnels for common types. All implementations are serializable.
 *
 * @author Dimitris Andreou
 * @since 11.0
 */
public final class Funnels {
  private Funnels() {}

  /**
   * Returns a funnel that encodes the characters of a {@code CharSequence} with the specified
   * {@code Charset}.
   *
   * @since 15.0
   */
  public static Funnel<CharSequence> stringFunnel(Charset charset) {
    return new StringCharsetFunnel(charset);
  }

  private static class StringCharsetFunnel implements Funnel<CharSequence>, Serializable {
    private final Charset charset;

    StringCharsetFunnel(Charset charset) {
      this.charset = charset;
    }

    @Override
    public void funnel(CharSequence from, PrimitiveSink into) {
      into.putString(from, charset);
    }

    @Override
    public String toString() {
      return "Funnels.stringFunnel(" + charset.name() + ")";
    }

    @Override
    public boolean equals( Object o) {
      if (o instanceof StringCharsetFunnel) {
        StringCharsetFunnel funnel = (StringCharsetFunnel) o;
        return this.charset.equals(funnel.charset);
      }
      return false;
    }

    @Override
    public int hashCode() {
      return StringCharsetFunnel.class.hashCode() ^ charset.hashCode();
    }

    Object writeReplace() {
      return new SerializedForm(charset);
    }

    private static class SerializedForm implements Serializable {
      private final String charsetCanonicalName;

      SerializedForm(Charset charset) {
        this.charsetCanonicalName = charset.name();
      }

      private Object readResolve() {
        return stringFunnel(Charset.forName(charsetCanonicalName));
      }

      private static final long serialVersionUID = 0;
    }
  }
}

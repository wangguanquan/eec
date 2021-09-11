/*
 * Copyright (c) 2017, guanquan.wang@yandex.com All Rights Reserved.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

package org.ttzero.excel.entity;

import java.util.Objects;

/**
 * @author guanquan.wang on 2017/9/26.
 */
public final class Tuple2<V1, V2> {
    public final V1 v1;
    public final V2 v2;

    public static <V1, V2> Tuple2<V1, V2> of(V1 v1, V2 v2) {
        return new Tuple2<>(v1, v2);
    }

    public Tuple2(V1 v1, V2 v2) {
        this.v1 = v1;
        this.v2 = v2;
    }

    public V1 v1() {
        return v1;
    }

    public V2 v2() {
        return v2;
    }

    public String stringV1() {
        return v1 != null ? v1.toString() : null;
    }

    public String stringV2() {
        return v2 != null ? v2.toString() : null;
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (!(o instanceof Tuple2)) return false;

        Tuple2 other = (Tuple2) o;

        return Objects.equals(v1, other.v1) && Objects.equals(v2, other.v2);
    }

    @Override
    public int hashCode() {
        return Objects.hashCode(v1) ^ Objects.hashCode(v2);
    }

    @Override
    public String toString() {
        return "Tuple2 { v1= " + v1 + ", v2=" + v2 + '}';
    }
}

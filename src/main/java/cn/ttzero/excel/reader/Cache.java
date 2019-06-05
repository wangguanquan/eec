/*
 * Copyright (c) 2019, guanquan.wang@yandex.com All Rights Reserved.
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

package cn.ttzero.excel.reader;

/**
 * Create by guanquan.wang at 2019-06-04 16:07
 */
public interface Cache<K, V> extends Iterable<Cache.Entry<K, V>> {

    /**
     * Returns the value to which the specified key is mapped,
     * or {@code null} if this cache contains no mapping for the key.
     *
     * @param k the key whose associated value is to be returned
     * @return the value to which the specified key is mapped, or
     *      {@code null} if this cache contains no mapping for the key
     */
    V get(K k);

    /**
     * Associates the specified value with the specified key in this cache.
     * If the cache previously contained a mapping for
     * the key, the old value is replaced by the specified value.
     *
     * @param k key with which the specified value is to be associated
     *          the key must not be null
     * @param v value to be associated with the specified key
     */
    void put(K k, V v);

    /**
     * Removes the mapping for a key from this cache if it is present
     * (optional operation).   More formally, if this cache contains a mapping
     * from key <tt>k</tt> to value <tt>v</tt>, that mapping
     * is removed.
     *
     * @param k key whose mapping is to be removed from the cache
     * @return the previous value associated with <tt>key</tt>, or
     *      <tt>null</tt> if there was no mapping for <tt>key</tt>.
     */
    V remove(K k);

    /**
     * Removes all of the mappings from this cache (optional operation).
     * The cache will be empty after this call returns.
     */
    void clear();

    /**
     * Returns the number of key-value mappings in this cache.
     *
     * @return the number of key-value mappings in this cache.
     */
    int size();

    /**
     * A map entry (key-value pair)
     */
    interface Entry<K, V> {
        /**
         * Returns the key corresponding to this entry.
         *
         * @return the key corresponding to this entry
         */
        K getKey();

        /**
         * Returns the value corresponding to this entry.  If the mapping
         * has been removed from the backing cache (by the iterator's
         * <tt>remove</tt> operation), the results of this call are undefined.
         *
         * @return the value corresponding to this entry
         */
        V getValue();
    }
}

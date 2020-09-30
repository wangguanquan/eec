/*
 * Copyright (c) 2017-2018, guanquan.wang@yandex.com All Rights Reserved.
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

package org.ttzero.excel.reader;

import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.function.Consumer;

/**
 * Implemented by the LRU page elimination algorithm.
 * The time complexity of the push, get, and remove is O(1).
 *
 * @author guanquan.wang at 2018-10-31 16:06
 */
public class FixSizeLRUCache<K, V> implements Cache<K, V> {
    private static class E<K, V> implements Cache.Entry<K, V> {
        private final K k;
        private V v;

        private E(K k, V v) {
            this.k = k;
            this.v = v;
        }

        @Override
        public K getKey() {
            return k;
        }
        @Override
        public V getValue() {
            return v;
        }
    }

    private static class Node<V> {
        private V data;
        private Node<V> prev, next;

        private Node(V data, Node<V> prev, Node<V> next) {
            this.data = data;
            this.prev = prev;
            this.next = next;
        }
    }

    /**
     * Double linked
     */
    private Node<E<K, V>> first, last;

    /**
     * The elements limit
     */
    private final int limit;

    /**
     * Size of elements
     */
    private int size;

    private final Map<K, Node<E<K, V>>> table;

    public static <K, V> FixSizeLRUCache<K, V> create() {
        return new FixSizeLRUCache<>();
    }

    public static <K, V> FixSizeLRUCache<K, V> create(int size) {
        return new FixSizeLRUCache<>(size);
    }

    /**
     * Create a fix size cache witch size is {@code 1 &lt;&lt; 9}
     */
    private FixSizeLRUCache() {
        this(1 << 9);
    }

    private FixSizeLRUCache(int limit) {
        this.limit = limit;
        // Create double limit size
        table = new HashMap<>(Math.round(limit * 1.25f));
    }

    /**
     * Returns the value to which the specified key is mapped,
     * or {@code null} if this cache contains no mapping for the key.
     *
     * @param k the key whose associated value is to be returned
     * @return the value to which the specified key is mapped, or
     *      {@code null} if this cache contains no mapping for the key
     */
    @Override
    public V get(K k) {
        final Node<E<K, V>> o;
        // Not found
        if (size == 0 || (o = table.get(k)) == null) return null;

        // Move node to head
        if (o != first) {
            if (o.next != null) {
                o.prev.next = o.next;
                o.next.prev = o.prev;
            } else {
                o.prev.next = null;
            }

            first.prev = o;
            o.next = first;
            o.prev = null;
            first = o;
        }

        return o.data.v;
    }

    /**
     * Associates the specified value with the specified key in this cache.
     * If the cache previously contained a mapping for
     * the key, the old value is replaced by the specified value.
     *
     * @param k key with which the specified value is to be associated
     *          the key must not be null
     * @param v value to be associated with the specified key
     */
    @Override
    public void put(K k, V v) {
        final Node<E<K, V>> o;
        // Insert at header if not found
        if (size == 0 || (o = table.get(k)) == null) {
            final Node<E<K, V>> f = first;
            final Node<E<K, V>> newNode = new Node<>(new E<>(k, v), null, f);
            first = newNode;
            if (f == null) last = newNode;
            else f.prev = newNode;

            table.put(k, newNode);

            if (size < limit) size++;
            else remove();
            // Replace the old value
        } else {
            o.data.v = v;
        }
    }

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
    @Override
    public V remove(K k) {
        final Node<E<K, V>> o;
        // Not found
        if (size == 0 || (o = table.get(k)) == null) return null;

        // Remove the keyword from the hash table to make them unsearchable
        table.remove(k);

        final Node<E<K, V>> prev = o.prev, next = o.next;

        if (prev == null) {
            first = next;
        } else {
            prev.next = next;
            o.prev = null;
        }

        if (next == null) {
            last = prev;
        } else {
            next.prev = prev;
            o.next = null;
        }

        if (--size == 0) {
            first = null;
        }

        V v = o.data.v;
        o.data = null;

        return v;
    }

    /**
     * Remove the last item
     *
     * @return the last item
     */
    protected E<K, V> remove() {
        final Node<E<K, V>> l = last;
        if (l == null) {
            return null;
        }
        final E<K, V> data = l.data;

        last = last.prev;
        if (last != null) {
            last.next = null;
        }

        table.remove(data.k);
        if (--size == 0) {
            first = null;
        }
        return data;
    }

    /**
     * Removes all of the mappings from this cache (optional operation).
     * The cache will be empty after this call returns.
     */
    @Override
    public void clear() {
        first = null;
        last = null;
        table.clear();
        size = 0;
    }

    /**
     * Returns the number of key-value mappings in this cache.
     *
     * @return the number of key-value mappings in this cache
     */
    @Override
    public int size() {
        return size;
    }

    /**
     * An inner iterator
     */
    private class CacheIterator implements Iterator<Entry<K, V>> {
        private Node<E<K, V>> first;

        private CacheIterator(Node<E<K, V>> first) {
            this.first = first;
        }

        @Override
        public boolean hasNext() {
            return first != null;
        }

        @Override
        public E<K, V> next() {
            E<K, V> e = first.data;
            first = first.next;
            return e;
        }
    }

    /**
     * Returns an iterator over elements of type {@code Cache.Entry<K, V>}.
     *
     * @return an Iterator.
     */
    @Override
    public Iterator<Entry<K, V>> iterator() {
        return new CacheIterator(first);
    }

    /**
     * Performs the given action for each element of the {@code Iterable}
     * until all elements have been processed or the action throws an
     * exception.  Unless otherwise specified by the implementing class,
     * actions are performed in the order of iteration (if an iteration order
     * is specified).  Exceptions thrown by the action are relayed to the
     * caller.
     *
     * <p>The default implementation behaves as if:
     * <pre>{@code
     *     for (T t : this)
     *         action.accept(t);
     * }</pre>
     *
     * @param action The action to be performed for each element
     * @throws NullPointerException if the specified action is null
     */
    @Override
    public void forEach(Consumer<? super Entry<K, V>> action) {
        Node<E<K, V>> f = first;
        for (; f != null; f = f.next) {
            action.accept(f.data);
        }
    }

}

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

import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.function.Consumer;

/**
 * 热词区
 * 由LRU页面淘汰算法实现push, get, remove方法时间复杂度均为O(1)
 * <p>
 * Create by guanquan.wang at 2018-10-31 16:06
 */
public class Hot<K,V> implements Iterable<Hot.E<K,V>> {
    public static class E<K,V> {
        private K k;
        private V v;

        private E(K k, V v) {
            this.k = k;
            this.v = v;
        }

        public K getK() {
            return k;
        }

        public V getV() {
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
     * double linked
     */
    private Node<E<K,V>> first, last;
    /**
     * elements limit
     * default size 64
     */
    private int limit;
    /**
     * size of elements
     */
    private int size;

    private Map<K, Node<E<K,V>>> table;

    public Hot() {
        this(1 << 9);
    }

    public Hot(int limit) {
        this.limit = limit;
        table = new HashMap<>(Math.round(limit * 1.25f));
    }

    /**
     * get by key
     *
     * @param k int
     * @return value
     */
    public V get(K k) {
        final Node<E<K,V>> o;
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
     * insert first
     *
     * @param v element
     */
    public void push(K k, V v) {
        final Node<E<K,V>> f = first;
        final Node<E<K,V>> newNode = new Node<>(new E<>(k, v), null, f);
        first = newNode;
        if (f == null) last = newNode;
        else f.prev = newNode;

        table.put(k, newNode);

        if (size < limit) size++;
        else remove();
    }

    /**
     * Remove the last value
     *
     * @return the last item
     */
    public E<K,V> remove() {
        final Node<E<K,V>> l = last;
        if (l == null) {
            return null;
        }
        final E<K,V> data = l.data;

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
     * Returns the cache word size
     * @return size of cache
     */
    public int size() {
        return size;
    }

    private class HotIterator implements Iterator<E<K,V>> {
        private Node<E<K,V>> first;
        private HotIterator(Node<E<K,V>> first) {
            this.first = first;
        }

        @Override
        public boolean hasNext() {
            return first != null;
        }

        @Override
        public E<K,V> next() {
            E<K,V> e = first.data;
            first = first.next;
            return e;
        }
    }

    @Override
    public Iterator<E<K,V>> iterator() {
        return new HotIterator(first);
    }

    @Override
    public void forEach(Consumer<? super E<K,V>> action) {
        Node<E<K,V>> f = first;
        for (; f != null; f = f.next) {
            action.accept(f.data);
        }
    }

}

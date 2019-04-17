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
import java.util.Map;

/**
 * 热词区
 * 由LRU页面淘汰算法实现push, get, remove方法时间复杂度均为O(1)
 *
 * Create by guanquan.wang at 2018-10-31 16:06
 */
public class Hot {
    private static class E {
        private int k;
        private String v;
        private E(int k, String v) {
            this.k = k;
            this.v = v;
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
    private Node<E> first, last;
    /**
     * elements limit
     * default size 64
     */
    private int limit;
    /**
     * size of elements
     */
    private int size;

    private Map<Integer, Node<E>> table;

    public Hot() {
        this(1 << 9);
    }

    public Hot(int limit) {
        this.limit = limit;
        table = new HashMap<>(Math.round(limit * 1.25f));
    }

    /**
     * get by key
     * @param k int
     * @return value
     */
    public String get(int k) {
        final Node<E> o;
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
     * @param v element
     */
    public void push(int k, String v) {
        final Node<E> f = first;
        final Node<E> newNode = new Node<>(new E(k, v), null, f);
        first = newNode;
        if (f == null) last = newNode;
        else f.prev = newNode;

        table.put(k, newNode);

        if (size < limit) size++;
        else remove();
    }

    /**
     * remove last
     * @return last item
     */
    public E remove() {
        final Node<E> l = last;
        final E data = l.data;

        last = last.prev;
        last.next = null;

        table.remove(data.k);
        return data;
    }
}

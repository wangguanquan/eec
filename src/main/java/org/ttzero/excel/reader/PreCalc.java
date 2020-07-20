/*
 * Copyright (c) 2017-2020, guanquan.wang@yandex.com All Rights Reserved.
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

import org.ttzero.excel.manager.Const;

import static org.ttzero.excel.entity.Sheet.int2Col;
import static org.ttzero.excel.reader.Row.toCellIndex;
import static org.ttzero.excel.reader.SharedStrings.toInt;

/**
 * Preprocessed calc node
 *
 * @author guanquan.wang at 2020-01-05 18:40
 */
class PreCalc {
    // Reference position
    private final long position;
    // calc string value
    private char[] cb;
    private Node head, tail;

    private static class Node {
        // From index, to index
        private int f, t;
        // Cell coordinate
        private long coordinate;
        // The next node
        private Node next;
    }

    PreCalc(long position) {
        this.position = position;
    }

    /* Pre-processing formula strings for fast getting */
    void setCalc(char[] cb) {
        this.cb = cb;

        int len = cb.length;
        int f = 0, m = f, n;
        for (; f < len; ) {
            for (; f < len && (cb[f] < 'A' || cb[f] > 'Z') && cb[f] != '"'; f++) ;
            // EOF
            if (f >= len) break;

            // Find the end quote tag
            if (cb[f] == '"') {
                for (; ++f < len && cb[f] != '"'; ) ;
                // EOF
                if (f >= len) break;
                f++;
                continue;
            } else n = f;

            // Column
            for (; f < len && cb[f] >= 'A' && cb[f] <= 'Z'; f++) ;

            int c = toCellIndex(cb, n, f);
            if (c < 0 || c > Const.Limit.MAX_COLUMNS_ON_SHEET) {
                continue;
            }

            int t = f; // tmp
            // Row
            for (; f < len && cb[f] >= '0' && cb[f] <= '9'; f++) ;
            if (f == t) continue;
            int r = toInt(cb, t, f);
            if (r < 0 || r > Const.Limit.MAX_ROWS_ON_SHEET) {
                continue;
            }

            Node node = new Node();
            node.f = m;
            node.t = n;

            if (head == null) {
                head = tail = node;
            } else {
                tail.next = node;
                tail = node;
            }

            Node next = new Node();
            next.coordinate = r << 14 | c;

            tail.next = next;
            tail = next;

            m = f;
        }

        if (m > 0 && m < f) {
            Node node = new Node();
            node.f = m;
            node.t = f;
            tail.next = node;
            tail = node;
        }
    }

    String get(long coordinate) {
        if (head == null) {
            return new String(cb, 0, cb.length);
        } else {
//            int t = (int) (position & 0x03);

            // Offset from first calc cell
            int offset_x = (int) ((coordinate & (1 << 14) - 1) - (position >> 28 & (1 << 14) - 1));
            int offset_y = (int) (((coordinate >> 14) & (1 << 20) - 1) - (position >> 42 & (1 << 20) - 1));


            Node node = head;
            StringBuilder buf = new StringBuilder();
            for (; node != null; ) {
                if (node.coordinate > 0) {
                    buf.append(int2Col((int) (node.coordinate & (1 << 14) - 1) + offset_x));
                    buf.append((int) (node.coordinate >> 14) + offset_y);
                } else {
                    buf.append(cb, node.f, node.t - node.f);
                }
                node = node.next;
            }

            return buf.toString();
        }
    }

    @Override
    public String toString() {
        return get((position >> 42 & (1 << 20) - 1) << 14 | ((position >> 28) & (1 << 14) - 1));
    }
}

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

package cn.ttzero.excel.entity.e7;

import java.nio.file.Path;
import java.util.Arrays;

/**
 * Created by guanquan.wang at 2018-02-26 14:00
 */
public class EmbedTemplate extends AbstractTemplate {

    public EmbedTemplate(Path zipPath, Workbook wb) {
        super(zipPath, wb);
    }

    @Override
    protected boolean isPlaceholder(String txt) {
        int n = txt.indexOf('$'), len = txt.length();
        if (n == -1 || n == len - 1) {
            return false;
        }
        boolean has = false;
        do {
            if (txt.charAt(n + 1) == '{') {
                int m = txt.indexOf('}', n + 3); // ${} key.length > 0
                if (m < n) {
                    break;
                }
                has = true;
                break;
            }
            n = txt.indexOf('$', n + 1);
        } while (n > -1 && n < len);
        return has;
    }

    @Override
    protected String getValue(String txt) {
        char[] values = txt.toCharArray();
        IntArray array = new IntArray();
        int n = txt.indexOf('$'), len = values.length;
        do {
            if (values[n + 1] == '{') {
                int m = txt.indexOf('}', n + 3); // ${} key.length > 0
                if (m < n) {
                    break;
                }
                array.add(n, m);
                n = m;
            }
            n = txt.indexOf('$', n + 1);
        } while (n > -1 && n < len);

        StringBuilder buf = new StringBuilder();
        for (int i = 0, size = array.size(), offset = 0; i < size; i++) {
            int[] idx = array.get(i);
            String key = new String(values, idx[0] + 2, idx[1] - idx[0] - 2).trim();
            if (map.containsKey(key)) {
                buf.append(values, offset, idx[0] - offset).append(map.get(key));
            } else {
                buf.append(values, offset, idx[1] - offset + 1);
            }
            offset = idx[1] + 1;
        }
        return buf.toString();
    }

    private final static class IntArray {
        private int[] elements;
        private int size;

        private IntArray() {
            elements = new int[8];
        }

        private int add(int i, int j) {
            final int length = elements.length, current = size;
            if (size + 1 >= length) {
                elements = Arrays.copyOf(elements, length << 1);
            }
            elements[size++] = i;
            elements[size++] = j;
            return current + 1;
        }

        private int size() {
            return size >> 1;
        }

        private boolean isEmpty() {
            return size == 0;
        }

        private int[] get(int index) {
            return new int[]{ elements[index <<= 1], elements[index + 1] };
        }

    }
}

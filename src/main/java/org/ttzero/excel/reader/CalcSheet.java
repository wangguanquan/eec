/*
 * Copyright (c) 2017-2019, guanquan.wang@yandex.com All Rights Reserved.
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

import org.dom4j.DocumentException;
import org.dom4j.Element;
import org.dom4j.io.SAXReader;

import java.io.InputStream;
import java.util.Arrays;
import java.util.Iterator;

import static org.ttzero.excel.reader.ExcelReader.LOGGER;
import static org.ttzero.excel.reader.ExcelReader.coordinateToLong;
import static org.ttzero.excel.reader.SharedStrings.toInt;
import static org.ttzero.excel.util.StringUtil.isNotEmpty;

/**
 * 支持解析公式的工作表，可以通过{@link #asCalcSheet}将普通工作表转为{@code CalcSheet}
 *
 * @author guanquan.wang at 2020-01-11 11:36
 * @deprecated 使用 {@link FullSheet}代替
 */
@Deprecated
public interface CalcSheet extends Sheet {

    /* Parse `calcChain` */
    static long[][] parseCalcChain(InputStream is) {
        SAXReader reader = SAXReader.createDefault();
        Element calcChain;
        try {
            calcChain = reader.read(is).getRootElement();
        } catch (DocumentException e) {
            LOGGER.warn("Part of `calcChain` has be damaged, It will be ignore all formulas.");
            return null;
        }

        Iterator<Element> ite = calcChain.elementIterator();
        int i = 1, n = 10;
        long[][] array = new long[n][];
        int[] indices = new int[n];
        for (; ite.hasNext(); ) {
            Element e = ite.next();
            // i: index of sheets
            // r: range
            String si = e.attributeValue("i"), r = e.attributeValue("r");
            if (isNotEmpty(si)) {
                i = toInt(si.toCharArray(), 0, si.length());
            }
            if (isNotEmpty(r)) {
                if (n < i) {
                    n <<= 1;
                    indices = Arrays.copyOf(indices, n);
                    long[][] _array = new long[n][];
                    for (int j = 0; j < n; j++) _array[j] = array[j]; // Do not copy hear.
                    array = _array;
                }
                long[] sub = array[i - 1];
                if (sub == null) {
                    sub = new long[10];
                    array[i - 1] = sub;
                }

                if (++indices[i - 1] > sub.length) {
                    long[] _sub = new long[sub.length << 1];
                    System.arraycopy(sub, 0, _sub, 0, sub.length);
                    array[i - 1] = sub = _sub;
                }
                sub[indices[i - 1] - 1] = coordinateToLong(r);
            }
        }

        i = 0;
        for (; i < n; i++) {
            if (indices[i] > 0) {
                long[] a = Arrays.copyOf(array[i], indices[i]);
                Arrays.sort(a);
                array[i] = a;
            } else array[i] = null;
        }
        return array;
    }

}

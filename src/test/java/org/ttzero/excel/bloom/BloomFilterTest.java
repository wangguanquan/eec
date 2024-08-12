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

package org.ttzero.excel.bloom;

import org.ttzero.excel.hash.StringBloomFilter;
import org.junit.Test;


import static org.junit.Assert.assertTrue;

/**
 * @author guanquan.wang at 2019-05-06 16:44
 */
public class BloomFilterTest {
    @Test public void test() {
        long start = System.currentTimeMillis();
        StringBloomFilter filter = StringBloomFilter.create(100000, 0.003);

        for (int index = 0; index < 100000; index++) {
            filter.put("abc_test_" + index);
        }
        int n = 0;
        for (int i = 0; i < 100000; i++) {
            if (filter.mightContain("abc_test_" + i)) {
                n++;
            }
        }
        System.out.println((System.currentTimeMillis() - start) + " " + n);
        assertTrue(n >= 99900);

    }
}

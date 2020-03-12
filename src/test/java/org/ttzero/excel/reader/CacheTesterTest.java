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

import org.junit.Test;

/**
 * @author guanquan.wang at 2020-03-11 15:39
 */
public class CacheTesterTest {
    @Test public void test() {
        Tester tester = new Tester.FixBinaryTester(1024);
        assert !tester.test(0);
        assert !tester.test(1);
        assert !tester.test(9);
        assert tester.test(1);
        assert tester.test(9);
        assert !tester.test(5);
        assert tester.test(0);

        for (int i = 0; i < 63; tester.test(i++)) ;

        assert !tester.test(1000);

        assert tester.size() == 16;

        assert tester.test(32);
        assert tester.test(1000);
        assert !tester.test(64);

        assert !tester.test(63);
        assert !tester.test(1024);

        assert tester.size() == 16;
        assert tester.limit() == 1087;

        assert tester.test(1024);
        assert !tester.test(1087);

        assert tester.size() == 16;
    }
}

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

package cn.ttzero.excel.entity;

import org.junit.Test;

import java.io.IOException;

/**
 * Create by guanquan.wang at 2019-05-07 17:41
 */
public class SharedStringsTest {
    @Test public void test() throws IOException {
        SharedStrings sst = new SharedStrings();
        int index = sst.get("abc");
        assert index == 1;

        index = sst.get("guanquan.wang");
        assert index == 2;

        index = sst.get("abc");
        assert index == 1;

        index = sst.get("guanquan.wang");
        assert index == 2;

        index = sst.get("guanquan.wang");
        assert index == 2;

        index = sst.get("test");
        assert index == 3;
    }
}

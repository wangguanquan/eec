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

import org.junit.Test;
import org.ttzero.excel.hash.StringBloomFilter;

import java.util.concurrent.CountDownLatch;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.TimeUnit;
import java.util.concurrent.atomic.AtomicInteger;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;

/**
 * @author guanquan.wang at 2019-05-06 16:44
 */
public class BloomFilterTest {
    @Test public void testStringFilter() {
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
        assertTrue(n >= 99997);
    }

    /**
     * 并发 put/mightContain 测试：验证 Strategy 内部 ThreadLocal<Murmur3_128Hasher>
     * 修复后不再出现哈希状态竞态。
     *
     * <p>修复前 Strategy 持有共享的 Murmur3_128Hasher 实例，
     * clear()→putBytes()→hash() 链条中 h1/h2/length/buffer 会被并发覆盖，
     * 导致 BloomFilter 位索引计算错误，mightContain() 对已 put 的值返回 false。</p>
     */
    @Test
    public void testConcurrentPutAndQuery() throws Exception {
        final int totalItems = 50_000;
        final int threads = 16;
        final int itemsPerThread = totalItems / threads;
        // 预期容量足够大，降低误判率干扰
        final StringBloomFilter filter = StringBloomFilter.create(totalItems * 2, 0.001);

        final ExecutorService executor = Executors.newFixedThreadPool(threads);
        final CountDownLatch startLatch = new CountDownLatch(1);
        final CountDownLatch putDoneLatch = new CountDownLatch(threads);
        final AtomicInteger putErrors = new AtomicInteger(0);

        // Phase 1: 并发 put
        for (int t = 0; t < threads; t++) {
            final int threadId = t;
            executor.execute(new Runnable() {
                @Override
                public void run() {
                    try {
                        startLatch.await();
                        int base = threadId * itemsPerThread;
                        for (int i = 0; i < itemsPerThread; i++) {
                            filter.put("item-" + (base + i));
                        }
                    } catch (Exception e) {
                        putErrors.incrementAndGet();
                    } finally {
                        putDoneLatch.countDown();
                    }
                }
            });
        }

        startLatch.countDown();
        putDoneLatch.await(60, TimeUnit.SECONDS);
        assertEquals("并发 put 阶段出现异常", 0, putErrors.get());

        // Phase 2: 并发 mightContain 验证（BloomFilter 不允许假阴性）
        final CountDownLatch queryDoneLatch = new CountDownLatch(threads);
        final AtomicInteger missedCount = new AtomicInteger(0);

        for (int t = 0; t < threads; t++) {
            final int threadId = t;
            executor.execute(new Runnable() {
                @Override
                public void run() {
                    try {
                        int base = threadId * itemsPerThread;
                        for (int i = 0; i < itemsPerThread; i++) {
                            if (!filter.mightContain("item-" + (base + i))) {
                                missedCount.incrementAndGet();
                            }
                        }
                    } finally {
                        queryDoneLatch.countDown();
                    }
                }
            });
        }

        queryDoneLatch.await(60, TimeUnit.SECONDS);
        executor.shutdown();

        // BloomFilter 的核心保证：put 过的元素 mightContain 必须返回 true（零假阴性）
        // 如果哈希器并发损坏，missedCount 会远大于 0
        assertEquals("BloomFilter 出现假阴性/并发哈希损坏：missed=" + missedCount.get(),
            0, missedCount.get());
    }
}

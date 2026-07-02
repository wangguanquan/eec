/*
 * Copyright (c) 2017-2022, guanquan.wang@yandex.com All Rights Reserved.
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

package org.ttzero.excel.entity;

import org.junit.Test;
import org.ttzero.excel.annotation.ExcelColumn;
import org.ttzero.excel.reader.ExcelReader;
import org.ttzero.excel.util.ExtBufferedWriter;

import java.nio.file.Files;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.CountDownLatch;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.TimeUnit;
import java.util.concurrent.atomic.AtomicReference;

import static org.junit.Assert.*;

/**
 * 并发安全测试：验证修复后的共享可变状态问题不再出现
 *
 * <p>覆盖以下修复项：</p>
 * <ul>
 *   <li>{@link Sheet#int2Col(int)} — 移除 {@code tmpBuf} 共享 char[][] 数组</li>
 *   <li>{@link ExtBufferedWriter#toChars(int)} / {@link ExtBufferedWriter#toChars(long)}
 *       — 移除 {@code CACHE_CHAR_ARRAY} 共享缓冲区</li>
 *   <li>并发 Workbook 导出 — 多线程同时调用 {@code writeData()} PUSH 模式写入</li>
 * </ul>
 *
 * @author onceMirsery (cmiracle@163.com)
 */
public class ConcurrencySafetyTest extends WorkbookTest {

    // ========== 1. Sheet.int2Col() 并发测试 ==========

    /**
     * 验证 int2Col() 在并发场景下返回正确的列字母。
     * 修复前使用共享 tmpBuf 数组，并发修改会导致列字母错乱。
     */
    @Test
    public void testInt2ColConcurrency() throws Exception {
        // 预先计算期望值
        final int[][] testCases = {
            {1}, {2}, {26},       // 单字母: A, B, Z
            {27}, {100}, {702},   // 双字母: AA, CV, ZZ
            {703}, {1000}, {16384} // 三字母: AAA, ALL, XFD
        };
        final String[] expected = {
            "A", "B", "Z",
            "AA", "CV", "ZZ",
            "AAA", "ALL", "XFD"
        };

        final int threads = 16;
        final int iterationsPerThread = 10_000;
        final ExecutorService executor = Executors.newFixedThreadPool(threads);
        final CountDownLatch startLatch = new CountDownLatch(1);
        final CountDownLatch doneLatch = new CountDownLatch(threads);
        final AtomicReference<String> failure = new AtomicReference<>(null);

        for (int t = 0; t < threads; t++) {
            executor.execute(new Runnable() {
                @Override
                public void run() {
                    try {
                        startLatch.await(); // 所有线程同时启动
                        for (int i = 0; i < iterationsPerThread; i++) {
                            for (int j = 0; j < testCases.length; j++) {
                                char[] result = Sheet.int2Col(testCases[j][0]);
                                String actual = new String(result);
                                if (!expected[j].equals(actual)) {
                                    failure.compareAndSet(null,
                                        "int2Col(" + testCases[j][0] + ") expected '"
                                        + expected[j] + "' but got '" + actual + "'");
                                    return;
                                }
                            }
                        }
                    } catch (Exception e) {
                        failure.compareAndSet(null, e.getMessage());
                    } finally {
                        doneLatch.countDown();
                    }
                }
            });
        }

        startLatch.countDown(); // 同时释放所有线程
        doneLatch.await(60, TimeUnit.SECONDS);
        executor.shutdown();

        assertNull("并发 int2Col() 结果错误: " + failure.get(), failure.get());
    }

    /**
     * 验证 toCoordinate() 在并发场景下返回正确的单元格地址。
     */
    @Test
    public void testToCoordinateConcurrency() throws Exception {
        final int threads = 16;
        final int iterations = 5_000;
        final ExecutorService executor = Executors.newFixedThreadPool(threads);
        final CountDownLatch startLatch = new CountDownLatch(1);
        final CountDownLatch doneLatch = new CountDownLatch(threads);
        final AtomicReference<String> failure = new AtomicReference<>(null);

        for (int t = 0; t < threads; t++) {
            final int threadId = t;
            executor.execute(new Runnable() {
                @Override
                public void run() {
                    try {
                        startLatch.await();
                        for (int i = 0; i < iterations; i++) {
                            // 每个线程用不同的列号范围
                            int col = (threadId * 100 + i % 100) + 1;
                            int row = (i % 5000) + 1;
                            String result = Sheet.toCoordinate(row, col);
                            // 验证格式：必须以字母开头、数字结尾
                            if (result == null || result.isEmpty()) {
                                failure.compareAndSet(null, "toCoordinate(" + row + "," + col + ") returned null/empty");
                                return;
                            }
                            // 验证字母部分
                            char[] colChars = Sheet.int2Col(col);
                            String expectedPrefix = new String(colChars);
                            if (!result.startsWith(expectedPrefix)) {
                                failure.compareAndSet(null,
                                    "toCoordinate(" + row + "," + col + ") expected prefix '"
                                    + expectedPrefix + "' but got '" + result + "'");
                                return;
                            }
                            // 验证数字部分
                            String numPart = result.substring(expectedPrefix.length());
                            if (!numPart.equals(String.valueOf(row))) {
                                failure.compareAndSet(null,
                                    "toCoordinate(" + row + "," + col + ") expected row "
                                    + row + " but got " + numPart);
                                return;
                            }
                        }
                    } catch (Exception e) {
                        failure.compareAndSet(null, e.getMessage());
                    } finally {
                        doneLatch.countDown();
                    }
                }
            });
        }

        startLatch.countDown();
        doneLatch.await(60, TimeUnit.SECONDS);
        executor.shutdown();

        assertNull("并发 toCoordinate() 结果错误: " + failure.get(), failure.get());
    }

    // ========== 2. ExtBufferedWriter.toChars() 并发测试 ==========

    /**
     * 验证 toChars(int) 在并发场景下返回正确的字符数组。
     * 修复前使用共享 CACHE_CHAR_ARRAY，并发调用会互相覆盖结果。
     */
    @Test
    public void testToCharsIntConcurrency() throws Exception {
        // 不同位数的测试值
        final int[] testValues = {
            0, 1, 9,               // 1位
            10, 99,                // 2位
            100, 999,              // 3位
            1000, 9999,            // 4位
            100000, 999999,        // 5-6位
            10000000, 99999999,    // 7-8位
            Integer.MAX_VALUE, Integer.MIN_VALUE // 极值
        };

        final String[] expected = new String[testValues.length];
        for (int i = 0; i < testValues.length; i++) {
            expected[i] = String.valueOf(testValues[i]);
        }

        final int threads = 16;
        final int iterations = 10_000;
        final ExecutorService executor = Executors.newFixedThreadPool(threads);
        final CountDownLatch startLatch = new CountDownLatch(1);
        final CountDownLatch doneLatch = new CountDownLatch(threads);
        final AtomicReference<String> failure = new AtomicReference<>(null);

        for (int t = 0; t < threads; t++) {
            executor.execute(new Runnable() {
                @Override
                public void run() {
                    try {
                        startLatch.await();
                        for (int i = 0; i < iterations; i++) {
                            for (int j = 0; j < testValues.length; j++) {
                                char[] result = ExtBufferedWriter.toChars(testValues[j]);
                                String actual = new String(result);
                                if (!expected[j].equals(actual)) {
                                    failure.compareAndSet(null,
                                        "toChars(int " + testValues[j] + ") expected '"
                                        + expected[j] + "' but got '" + actual + "'");
                                    return;
                                }
                            }
                        }
                    } catch (Exception e) {
                        failure.compareAndSet(null, e.getMessage());
                    } finally {
                        doneLatch.countDown();
                    }
                }
            });
        }

        startLatch.countDown();
        doneLatch.await(60, TimeUnit.SECONDS);
        executor.shutdown();

        assertNull("并发 toChars(int) 结果错误: " + failure.get(), failure.get());
    }

    /**
     * 验证 toChars(long) 在并发场景下返回正确的字符数组。
     */
    @Test
    public void testToCharsLongConcurrency() throws Exception {
        final long[] testValues = {
            0L, 1L, 9L,
            100L, 9999L,
            1000000L, 99999999999L,
            Long.MAX_VALUE, Long.MIN_VALUE
        };

        final String[] expected = new String[testValues.length];
        for (int i = 0; i < testValues.length; i++) {
            expected[i] = String.valueOf(testValues[i]);
        }

        final int threads = 16;
        final int iterations = 10_000;
        final ExecutorService executor = Executors.newFixedThreadPool(threads);
        final CountDownLatch startLatch = new CountDownLatch(1);
        final CountDownLatch doneLatch = new CountDownLatch(threads);
        final AtomicReference<String> failure = new AtomicReference<>(null);

        for (int t = 0; t < threads; t++) {
            executor.execute(new Runnable() {
                @Override
                public void run() {
                    try {
                        startLatch.await();
                        for (int i = 0; i < iterations; i++) {
                            for (int j = 0; j < testValues.length; j++) {
                                char[] result = ExtBufferedWriter.toChars(testValues[j]);
                                String actual = new String(result);
                                if (!expected[j].equals(actual)) {
                                    failure.compareAndSet(null,
                                        "toChars(long " + testValues[j] + ") expected '"
                                        + expected[j] + "' but got '" + actual + "'");
                                    return;
                                }
                            }
                        }
                    } catch (Exception e) {
                        failure.compareAndSet(null, e.getMessage());
                    } finally {
                        doneLatch.countDown();
                    }
                }
            });
        }

        startLatch.countDown();
        doneLatch.await(60, TimeUnit.SECONDS);
        executor.shutdown();

        assertNull("并发 toChars(long) 结果错误: " + failure.get(), failure.get());
    }

    /**
     * 验证 toChars 返回的是独立数组——修改一个返回值不应影响另一个。
     * 这是修复 CACHE_CHAR_ARRAY 共享问题的核心验证。
     */
    @Test
    public void testToCharsReturnsIndependentArray() {
        char[] a = ExtBufferedWriter.toChars(123);
        char[] b = ExtBufferedWriter.toChars(456);
        // 修改 a 不应影响 b
        a[0] = 'X';
        assertEquals("456", new String(b));
        assertEquals("X23", new String(a));
    }

    /**
     * 验证 int2Col 返回的是独立数组——修改一个返回值不应影响另一个。
     * 这是修复 tmpBuf 共享问题的核心验证。
     */
    @Test
    public void testInt2ColReturnsIndependentArray() {
        char[] a = Sheet.int2Col(1);  // A
        char[] b = Sheet.int2Col(2);  // B
        // 修改 a 不应影响 b
        a[0] = 'X';
        assertEquals("B", new String(b));
        assertEquals("X", new String(a));
    }

    // ========== 3. 并发 Workbook 导出集成测试 ==========

    /**
     * 测试实体类
     */
    public static class ExportItem {
        @ExcelColumn("ID")
        private int id;
        @ExcelColumn("名称")
        private String name;
        @ExcelColumn("值")
        private double value;

        public ExportItem() { }

        public ExportItem(int id, String name, double value) {
            this.id = id;
            this.name = name;
            this.value = value;
        }

        public int getId() { return id; }
        public String getName() { return name; }
        public double getValue() { return value; }
    }

    /**
     * 并发 Workbook 导出集成测试：
     * 多个线程各自创建独立的 Workbook，使用 writeData() PUSH 模式分批写入数据，
     * 最终验证每个文件的数据行数和完整性。
     *
     * <p>此测试覆盖了 {@code XMLWorksheetWriter.getColumn()}、{@code Sheet.int2Col()}
     * 和 {@code ExtBufferedWriter.toChars()} 在并发场景下的综合表现。</p>
     */
    @Test
    public void testConcurrentWorkbookExport() throws Exception {
        getOutputTestPath();
        final int taskCount = 8;
        final int pagesPerTask = 3;
        final int rowsPerPage = 100;
        final int expectedRows = pagesPerTask * rowsPerPage;

        final ExecutorService executor = Executors.newFixedThreadPool(taskCount);
        final CountDownLatch startLatch = new CountDownLatch(1);
        final CountDownLatch doneLatch = new CountDownLatch(taskCount);
        final AtomicReference<String> failure = new AtomicReference<>(null);

        for (int t = 0; t < taskCount; t++) {
            final int taskId = t;
            executor.execute(new Runnable() {
                @Override
                public void run() {
                    try {
                        startLatch.await();
                        String fileName = "concurrent_task_" + taskId + ".xlsx";
                        java.nio.file.Path filePath = defaultTestPath.resolve(fileName);

                        // 创建独立的 Workbook
                        Workbook workbook = new Workbook("Task-" + taskId);
                        ListSheet<ExportItem> sheet = new ListSheet<>("数据");
                        workbook.addSheet(sheet);

                        // 分批写入数据 (PUSH 模式)
                        for (int page = 0; page < pagesPerTask; page++) {
                            List<ExportItem> pageData = new ArrayList<>(rowsPerPage);
                            int baseId = taskId * 100000 + page * rowsPerPage;
                            for (int i = 0; i < rowsPerPage; i++) {
                                pageData.add(new ExportItem(
                                    baseId + i,
                                    "Task" + taskId + "-Item" + i,
                                    taskId * 100.0 + i
                                ));
                            }
                            sheet.writeData(pageData);
                        }

                        workbook.writeTo(filePath);

                        // 验证输出文件
                        try (ExcelReader reader = ExcelReader.read(filePath)) {
                            long actualRows = reader.sheet(0).dataRows().count();
                            if (actualRows != expectedRows) {
                                failure.compareAndSet(null,
                                    "Task-" + taskId + ": expected " + expectedRows
                                    + " rows but got " + actualRows);
                            }
                        }

                        // 清理测试文件
                        Files.deleteIfExists(filePath);
                    } catch (Exception e) {
                        failure.compareAndSet(null, "Task-" + taskId + ": " + e.getMessage());
                    } finally {
                        doneLatch.countDown();
                    }
                }
            });
        }

        startLatch.countDown();
        doneLatch.await(120, TimeUnit.SECONDS);
        executor.shutdown();

        assertNull("并发导出测试失败: " + failure.get(), failure.get());
    }

    /**
     * 并发多 Sheet 导出测试：
     * 每个任务创建一个 Workbook 包含 3 个 Sheet，分别写入不同数据，
     * 验证各 Sheet 的数据不会交叉污染。
     */
    @Test
    public void testConcurrentMultiSheetExport() throws Exception {
        getOutputTestPath();
        final int taskCount = 6;
        final int rowsPerSheet = 200;

        final ExecutorService executor = Executors.newFixedThreadPool(taskCount);
        final CountDownLatch startLatch = new CountDownLatch(1);
        final CountDownLatch doneLatch = new CountDownLatch(taskCount);
        final AtomicReference<String> failure = new AtomicReference<>(null);

        for (int t = 0; t < taskCount; t++) {
            final int taskId = t;
            executor.execute(new Runnable() {
                @Override
                public void run() {
                    try {
                        startLatch.await();
                        String fileName = "concurrent_multi_sheet_" + taskId + ".xlsx";
                        java.nio.file.Path filePath = defaultTestPath.resolve(fileName);

                        Workbook workbook = new Workbook("MultiSheet-" + taskId);
                        ListSheet<ExportItem> sheet1 = new ListSheet<>("Sheet-A");
                        ListSheet<ExportItem> sheet2 = new ListSheet<>("Sheet-B");
                        ListSheet<ExportItem> sheet3 = new ListSheet<>("Sheet-C");
                        workbook.addSheet(sheet1).addSheet(sheet2).addSheet(sheet3);

                        // 每个 Sheet 写入不同标记的数据
                        List<ExportItem> dataA = new ArrayList<>(rowsPerSheet);
                        List<ExportItem> dataB = new ArrayList<>(rowsPerSheet);
                        List<ExportItem> dataC = new ArrayList<>(rowsPerSheet);
                        for (int i = 0; i < rowsPerSheet; i++) {
                            dataA.add(new ExportItem(taskId * 1000 + i, "A-" + taskId + "-" + i, 1.0));
                            dataB.add(new ExportItem(taskId * 2000 + i, "B-" + taskId + "-" + i, 2.0));
                            dataC.add(new ExportItem(taskId * 3000 + i, "C-" + taskId + "-" + i, 3.0));
                        }
                        sheet1.writeData(dataA);
                        sheet2.writeData(dataB);
                        sheet3.writeData(dataC);

                        workbook.writeTo(filePath);

                        // 验证每个 Sheet
                        try (ExcelReader reader = ExcelReader.read(filePath)) {
                            assertEquals("Sheet count mismatch", 3, reader.getSheetCount());
                            for (int s = 0; s < 3; s++) {
                                long count = reader.sheet(s).dataRows().count();
                                if (count != rowsPerSheet) {
                                    failure.compareAndSet(null,
                                        "Task-" + taskId + " Sheet-" + s + ": expected "
                                        + rowsPerSheet + " rows but got " + count);
                                }
                            }
                            // 验证 Sheet-A 的第一行数据名称前缀
                            String firstName = reader.sheet(0).dataRows()
                                .findFirst().map(r -> r.getString(1)).orElse(null);
                            assertNotNull("Sheet-A first row name is null", firstName);
                            if (!firstName.startsWith("A-" + taskId + "-")) {
                                failure.compareAndSet(null,
                                    "Task-" + taskId + " Sheet-A data polluted: got '"
                                    + firstName + "'");
                            }
                        }

                        Files.deleteIfExists(filePath);
                    } catch (Exception e) {
                        failure.compareAndSet(null, "Task-" + taskId + ": " + e.getMessage());
                    } finally {
                        doneLatch.countDown();
                    }
                }
            });
        }

        startLatch.countDown();
        doneLatch.await(120, TimeUnit.SECONDS);
        executor.shutdown();

        assertNull("并发多Sheet导出失败: " + failure.get(), failure.get());
    }
}

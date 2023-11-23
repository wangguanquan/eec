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
import org.ttzero.excel.entity.style.Fill;
import org.ttzero.excel.entity.style.Font;
import org.ttzero.excel.entity.style.Horizontals;
import org.ttzero.excel.entity.style.PatternType;
import org.ttzero.excel.entity.style.Styles;
import org.ttzero.excel.manager.Const;
import org.ttzero.excel.processor.StyleProcessor;
import org.ttzero.excel.reader.Dimension;
import org.ttzero.excel.reader.ExcelReader;
import org.ttzero.excel.reader.MergeSheet;
import org.ttzero.excel.reader.Row;

import java.awt.Color;
import java.io.IOException;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

/**
 * @author guanquan.wang at 2022-07-30 22:31
 */
public class ReportDesignTest extends WorkbookTest {

    @Test public void testMergedCells() throws IOException {
        List<E> list = testData();
        new Workbook().cancelZebraLine().setAutoSize(true)
            .addSheet(new ListSheet<>(list, createColumns()).setStyleProcessor(new GroupStyleProcessor<>()).hideGridLines())
            .writeTo(defaultTestPath.resolve("Group Style Processor.xlsx"));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("Group Style Processor.xlsx"))) {
            List<Map<String, Object>> expectList = reader.sheet(0).dataRows().map(Row::toMap).collect(Collectors.toList());

            assert expectList.size() == list.size();
            for (int i = 0, len = expectList.size(); i < len; i++) {
                Map<String, Object> m = expectList.get(i);
                E e = list.get(i);
                assert m.get("日期").equals(e.date);
                assert m.get("客户名称").equals(e.customer);
                assert m.get("商品名称").equals(e.productName);
                assert m.get("品牌").equals(e.brand);
                assert m.get("单位").equals(e.unit);
                assert m.get("数量").equals(e.num);
                assert new BigDecimal(m.get("含税单价").toString()).setScale(2, BigDecimal.ROUND_HALF_DOWN).equals(e.unitPrice.setScale(2, BigDecimal.ROUND_HALF_DOWN));
                assert new BigDecimal(m.get("含税总额").toString()).setScale(2, BigDecimal.ROUND_HALF_DOWN).equals(e.totalAmount.setScale(2, BigDecimal.ROUND_HALF_DOWN));
                assert m.get("出库数量").equals(e.outNum);
                assert m.get("关联订单").equals(e.orderNo);
            }

            Iterator<Row> iter = reader.sheet(0).iterator();
            assert iter.hasNext();
            org.ttzero.excel.reader.Row header = iter.next();
            assert "日期".equals(header.getString(0));
            assert "客户名称".equals(header.getString(1));
            assert "商品名称".equals(header.getString(2));
            assert "品牌".equals(header.getString(3));
            assert "单位".equals(header.getString(4));
            assert "数量".equals(header.getString(5));
            assert "含税单价".equals(header.getString(6));
            assert "含税总额".equals(header.getString(7));
            assert "出库数量".equals(header.getString(8));
            assert "关联订单".equals(header.getString(9));

            int r = 0;
            String orderNo = null;
            while (iter.hasNext()) {
                org.ttzero.excel.reader.Row row = iter.next();

                if (!row.getString("关联订单").equals(orderNo)) {
                    r++;
                    orderNo = row.getString("关联订单");
                }

                Styles styles = row.getStyles();

                assert styles.getHorizontal(row.getCellStyle(0)) == Horizontals.CENTER;
                assert styles.getHorizontal(row.getCellStyle(1)) == Horizontals.LEFT;
                assert styles.getHorizontal(row.getCellStyle(2)) == Horizontals.LEFT;
                assert styles.getHorizontal(row.getCellStyle(3)) == Horizontals.CENTER;
                assert styles.getHorizontal(row.getCellStyle(4)) == Horizontals.CENTER;
                assert styles.getHorizontal(row.getCellStyle(5)) == Horizontals.RIGHT;
                assert styles.getHorizontal(row.getCellStyle(6)) == Horizontals.RIGHT;
                assert styles.getHorizontal(row.getCellStyle(7)) == Horizontals.RIGHT;
                assert styles.getHorizontal(row.getCellStyle(8)) == Horizontals.RIGHT;
                assert styles.getHorizontal(row.getCellStyle(9)) == Horizontals.CENTER;


                int style = row.getCellStyle(0);
                Fill fill = styles.getFill(style);

                if ((r & 1) == 1) {
                    assert fill == null || fill.getPatternType() == PatternType.none;
                } else {
                    assert fill != null && fill.getPatternType() == PatternType.solid && fill.getFgColor().equals(new Color(233, 234, 236));
                }
            }
        }
    }

    @Test public void testMergedCells1() throws IOException {
        List<E> list = testData();
        // 用于保存合并单元格
        List<Dimension> mergeCells = new ArrayList<>();
        String date = null, order = null;
        int row = 2, dateFrom = row, orderFrom = row; // 记录订单/日期的起始位置
        E summary = null, allSummary = createSummary();
        for (int i = 0, size = list.size(); i < size; ) {
            E e = list.get(i);
            if (!e.orderNo.equals(order)) {
                if (order != null) {
                    list.add(i++, summary);
                    size++;
                    // 合并客户名和订单号
                    mergeCells.add(new Dimension(orderFrom, (short) 2, row, (short) 2));
                    mergeCells.add(new Dimension(orderFrom, (short) 10, row, (short) 10));
                    // 合并小计
                    mergeCells.add(new Dimension(row, (short) 3, row, (short) 5));
                    row++;
                }
                summary = createSummary();
                summary.orderNo = e.orderNo;
                summary.date = e.date;

                order = e.orderNo;
                orderFrom = row;
            } else {
                e.orderNo = null;
                e.customer = null;
            }
            if (!e.date.equals(date)) {
                if (date != null) {
                    // 合并日期
                    mergeCells.add(new Dimension(dateFrom, (short) 1, row - 1, (short) 1));
                }
                dateFrom = row;
                date = e.date;
            } else e.date = null;

            // 累计
            summary.num += e.num;
            summary.totalAmount = summary.totalAmount.add(e.totalAmount);

            allSummary.num += e.num;
            allSummary.totalAmount = allSummary.totalAmount.add(e.totalAmount);

            i++;
            row++;
        }
        // 添加最后一个订单小计以及合计数据
        list.add(summary);
        mergeCells.add(new Dimension(dateFrom, (short) 1, row, (short) 1));
        mergeCells.add(new Dimension(orderFrom, (short) 2, row, (short) 2));
        mergeCells.add(new Dimension(orderFrom, (short) 10, row, (short) 10));
        mergeCells.add(new Dimension(row, (short) 3, row, (short) 5));

        allSummary.date = "总计：";
        allSummary.productName = null;
        allSummary.orderNo = "--";
        list.add(allSummary);
        row++;
        mergeCells.add(new Dimension(row, (short) 1, row, (short) 5));

        new Workbook().cancelZebraLine().setAutoSize(true)
            .addSheet(new ListSheet<>(list, createColumns())
                .setStyleProcessor(new GroupStyleProcessor2<>())
                .putExtProp(Const.ExtendPropertyKey.MERGE_CELLS, mergeCells).hideGridLines()).writeTo(defaultTestPath.resolve("Report Design.xlsx"));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("Report Design.xlsx"))) {
            MergeSheet sheet = reader.sheet(0).asMergeSheet();
            List<Dimension> mergeCells1 = sheet.getMergeCells();

            // Expect merge cells
            assert mergeCells.size() == mergeCells1.size();
            for (int i = 0; i < mergeCells.size(); i++) {
                Dimension d0 = mergeCells.get(i), d1 = mergeCells1.get(i);
                assert d0.equals(d1);
            }

            List<Map<String, Object>> expectList = sheet.asSheet().dataRows().map(Row::toMap).collect(Collectors.toList());

            assert expectList.size() == list.size();
            for (int i = 0, len = expectList.size(); i < len; i++) {
                Map<String, Object> m = expectList.get(i);
                E e = list.get(i);
                assert e.date == null ? m.get("日期") == null || m.get("日期").toString().isEmpty() : m.get("日期").equals(e.date);
                assert e.customer == null ? m.get("客户名称") == null || m.get("客户名称").toString().isEmpty() : m.get("客户名称").equals(e.customer);
                assert e.productName == null ? m.get("商品名称") == null || m.get("商品名称").toString().isEmpty() : m.get("商品名称").equals(e.productName);
                assert e.brand == null ? m.get("品牌") == null || m.get("品牌").toString().isEmpty() : m.get("品牌").equals(e.brand);
                assert e.unit == null ? m.get("单位") == null || m.get("单位").toString().isEmpty() : m.get("单位").equals(e.unit);
                assert e.num == null ? m.get("数量") == null || m.get("数量").toString().isEmpty() : m.get("数量").equals(e.num);
                assert e.unitPrice == null ? m.get("含税单价") == null || m.get("含税单价").toString().isEmpty() : new BigDecimal(m.get("含税单价").toString()).setScale(2, BigDecimal.ROUND_HALF_DOWN).equals(e.unitPrice.setScale(2, BigDecimal.ROUND_HALF_DOWN));
                assert e.totalAmount == null ? m.get("含税总额") == null || m.get("含税总额").toString().isEmpty() : new BigDecimal(m.get("含税总额").toString()).setScale(2, BigDecimal.ROUND_HALF_DOWN).equals(e.totalAmount.setScale(2, BigDecimal.ROUND_HALF_DOWN));
                assert e.outNum == null ? m.get("出库数量") == null || m.get("出库数量").toString().isEmpty() : m.get("出库数量").equals(e.outNum);
                assert e.orderNo == null ? m.get("关联订单") == null || m.get("日期").toString().isEmpty() : m.get("关联订单").equals(e.orderNo);
            }

            Iterator<Row> iter = reader.sheet(0).iterator();
            assert iter.hasNext();
            org.ttzero.excel.reader.Row header = iter.next();
            assert "日期".equals(header.getString(0));
            assert "客户名称".equals(header.getString(1));
            assert "商品名称".equals(header.getString(2));
            assert "品牌".equals(header.getString(3));
            assert "单位".equals(header.getString(4));
            assert "数量".equals(header.getString(5));
            assert "含税单价".equals(header.getString(6));
            assert "含税总额".equals(header.getString(7));
            assert "出库数量".equals(header.getString(8));
            assert "关联订单".equals(header.getString(9));
        }
    }

    public static List<E> testData() {
        return new ArrayList<E>() {{
            add(new E("王先生", "纽仕兰新西兰进口牛奶3.5g蛋白质牧场草饲高钙礼盒全脂纯牛奶乳品250ml*24 整箱装", "纽仕兰", "箱", "X33322071291186", "2022-07-12", 2, 1, new BigDecimal("220.00"), new BigDecimal("440.00")));
            add(new E("王先生", "百医卫仕护必安 口罩N95口罩", "百医卫仕", "包", "X33322071291186", "2022-07-12", 1, 1, new BigDecimal("59.90"), new BigDecimal("59.90")));
            add(new E("张老板", "ABB 模块化按钮指示灯附件(一常开触点)；MCB-10", "ABB", "个", "X33322070700901", "2022-07-07", 1, 1, new BigDecimal("220.00"), new BigDecimal("220.00")));
            add(new E("张老板", "霍尼韦尔(Honeywell) 大带灯片(只用于交流)；AB22-D-AC220V-G", "霍尼韦尔", "个", "X33322070700901", "2022-07-07", 1, 1, new BigDecimal("111.00"), new BigDecimal("111.00")));
            add(new E("张先生", "牧田/MAKITA A-49579扭转十字批头 PH2长65mm六支装", "牧田", "支", "X33322070500539", "2022-07-05", 3, 1, new BigDecimal("200.00"), new BigDecimal("570.00")));
            add(new E("田女士", "日本进口 尤妮佳(unicharm)舒蔻雅致棉柔型化妆棉 66片（卸妆棉天然棉保湿柔软亲肤 水润呵护）", "尤妮佳", "包", "S33322070500458", "2022-07-05", 1, 1, new BigDecimal("24.12"), new BigDecimal("24.12")));
            add(new E("李老板", "德力西电气铜开口鼻1", "德力西电气", "只", "D33322062000190", "2022-06-20", 1, 1, new BigDecimal("33.00"), new BigDecimal("33.00")));
        }};
    }

    public static Column[] createColumns() {
        return new Column[]{
            new Column("日期", "date").setStyleProcessor((n, i, st) -> st.modifyHorizontal(i, Horizontals.CENTER))
            , new Column("客户名称", "customer")
            , new Column("商品名称", "productName").setWidth(30.68D).setWrapText(true)
            , new Column("品牌", "brand").setStyleProcessor((n, i, st) -> st.modifyHorizontal(i, Horizontals.CENTER))
            , new Column("单位", "unit").setStyleProcessor((n, i, st) -> st.modifyHorizontal(i, Horizontals.CENTER))
            , new Column("数量", "num")
            , new Column("含税单价", "unitPrice").setNumFmt("#,##0.00_);0_)")
            , new Column("含税总额", "totalAmount").setNumFmt("#,##0.00_);0_)")
            , new Column("出库数量", "outNum")
            , new Column("关联订单", "orderNo").setStyleProcessor((n, i, st) -> st.modifyHorizontal(i, Horizontals.CENTER))
        };
    }

    public static E createSummary() {
        E summary = new E() {
            @Override
            public boolean isSummary() {
                return true;
            }
        };
        summary.productName = "小计：";
        summary.num = 0;
        summary.outNum = 0;
        summary.totalAmount = BigDecimal.ZERO;

        return summary;
    }

    public static class E implements Group, Summary {
        String customer, productName, brand, unit, orderNo, date;
        Integer num, outNum;
        BigDecimal unitPrice, totalAmount;

        public E() {
        }

        public E(String customer, String productName, String brand, String unit, String orderNo, String date, Integer num, Integer outNum, BigDecimal unitPrice, BigDecimal totalAmount) {
            this.customer = customer;
            this.productName = productName;
            this.brand = brand;
            this.unit = unit;
            this.orderNo = orderNo;
            this.date = date;
            this.num = num;
            this.outNum = outNum;
            this.unitPrice = unitPrice;
            this.totalAmount = totalAmount;
        }

        @Override
        public String groupBy() {
            return orderNo;
        }

        @Override
        public boolean isSummary() {
            return false;
        }
    }

    // =======================公共部分=======================
    public interface Group {
        String groupBy();
    }

    public interface Summary {
        boolean isSummary();
    }

    public static class GroupStyleProcessor<U extends Group> implements StyleProcessor<U> {
        private String group;
        private int s, o;

        @Override
        public int build(U u, int style, Styles sst) {
            if (group == null) {
                group = u.groupBy();
                s = sst.addFill(new Fill(PatternType.solid, new Color(233, 234, 236)));
                return style;
            } else if (u.groupBy() != null && !group.equals(u.groupBy())) {
                group = u.groupBy();
                o ^= 1;
            }
            return o == 1 ? Styles.clearFill(style) | s : style;
        }
    }

    public static class GroupStyleProcessor2<U extends Group & Summary> implements StyleProcessor<U> {
        private String group;
        private int s, o, i;

        @Override
        public int build(U u, int style, Styles sst) {
            if (group == null) {
                group = u.groupBy();
                s = sst.addFill(new Fill(PatternType.solid, new Color(233, 234, 236)));
                return style;
            }
            // 小计加粗字体
            if (u.isSummary()) {
                Font font = sst.getFont(style).clone();
                font.bold();
                style = sst.modifyFont(style, font);
            } else if (u.groupBy() != null && !group.equals(u.groupBy())) {
                group = u.groupBy();
                o ^= 1;
                i = 0;
            }
            return o == 1 && ++i > 1 ? Styles.clearFill(style) | s : style;
        }
    }
}

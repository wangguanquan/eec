/*
 * Copyright (c) 2017-2023, guanquan.wang@yandex.com All Rights Reserved.
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
import org.ttzero.excel.entity.e7.XMLWorksheetWriter;
import org.ttzero.excel.entity.style.Border;
import org.ttzero.excel.entity.style.BorderStyle;
import org.ttzero.excel.entity.style.Fill;
import org.ttzero.excel.entity.style.Font;
import org.ttzero.excel.entity.style.Horizontals;
import org.ttzero.excel.entity.style.PatternType;
import org.ttzero.excel.entity.style.Styles;
import org.ttzero.excel.manager.Const;
import org.ttzero.excel.processor.Converter;
import org.ttzero.excel.processor.StyleProcessor;
import org.ttzero.excel.reader.Cell;
import org.ttzero.excel.reader.Dimension;
import org.ttzero.excel.reader.ExcelReader;
import org.ttzero.excel.reader.HeaderRow;
import org.ttzero.excel.reader.Sheet;
import org.ttzero.excel.util.StringUtil;

import java.awt.Color;
import java.io.IOException;
import java.math.BigDecimal;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Iterator;
import java.util.List;
import java.util.Objects;
import java.util.stream.Collectors;

import static org.ttzero.excel.reader.Cell.BLANK;
import static org.ttzero.excel.reader.Cell.BOOL;
import static org.ttzero.excel.reader.Cell.CHARACTER;
import static org.ttzero.excel.reader.Cell.DATE;
import static org.ttzero.excel.reader.Cell.DATETIME;
import static org.ttzero.excel.reader.Cell.DECIMAL;
import static org.ttzero.excel.reader.Cell.DOUBLE;
import static org.ttzero.excel.reader.Cell.INLINESTR;
import static org.ttzero.excel.reader.Cell.LONG;
import static org.ttzero.excel.reader.Cell.NUMERIC;
import static org.ttzero.excel.reader.Cell.SST;
import static org.ttzero.excel.reader.Cell.TIME;

/**
 * @author guanquan.wang at 2023-04-04 22:38
 */
public class ListObjectSheetTest2 extends WorkbookTest {
    @Test public void testSpecifyRowWrite() throws IOException {
        List<ListObjectSheetTest.Item> list = ListObjectSheetTest.Item.randomTestData();
        new Workbook().setAutoSize(true)
            .addSheet(new ListSheet<>(list).setStartRowIndex(4))
            .writeTo(defaultTestPath.resolve("test specify row 4 ListSheet.xlsx"));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("test specify row 4 ListSheet.xlsx"))) {
            List<ListObjectSheetTest.Item> readList = reader.sheet(0).header(4).rows().map(row -> row.to(ListObjectSheetTest.Item.class)).collect(Collectors.toList());
            assert list.size() == readList.size();
            for (int i = 0, len = list.size(); i < len; i++)
                assert list.get(i).equals(readList.get(i));
        }
    }

    @Test public void testSpecifyRowStayA1Write() throws IOException {
        List<ListObjectSheetTest.Item> list = ListObjectSheetTest.Item.randomTestData();
        new Workbook().setAutoSize(true)
            .addSheet(new ListSheet<>(list).setStartRowIndex(4, false))
            .writeTo(defaultTestPath.resolve("test specify row 4 stay A1 ListSheet.xlsx"));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("test specify row 4 stay A1 ListSheet.xlsx"))) {
            List<ListObjectSheetTest.Item> readList = reader.sheet(0).bind(ListObjectSheetTest.Item.class, 4).rows().map(row -> (ListObjectSheetTest.Item) row.get()).collect(Collectors.toList());
            assert list.size() == readList.size();
            for (int i = 0, len = list.size(); i < len; i++)
                assert list.get(i).equals(readList.get(i));
        }
    }

    @Test public void testSpecifyRowAndColWrite() throws IOException {
        List<ListObjectSheetTest.Item> list = ListObjectSheetTest.Item.randomTestData(10);
        new Workbook().setAutoSize(true)
            .addSheet(new ListSheet<ListObjectSheetTest.Item>("Item"
                , new Column("id").setColIndex(3)
                , new Column("name").setColIndex(4))
                .setData(list)
                .setStartRowIndex(4)
            ).writeTo(defaultTestPath.resolve("test specify row and cel ListSheet.xlsx"));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("test specify row and cel ListSheet.xlsx"))) {
            List<ListObjectSheetTest.Item> readList = reader.sheet(0).bind(ListObjectSheetTest.Item.class, 4).rows().map(row -> (ListObjectSheetTest.Item) row.get()).collect(Collectors.toList());
            assert list.size() == readList.size();
            for (int i = 0, len = list.size(); i < len; i++)
                assert list.get(i).equals(readList.get(i));
        }
    }

    @Test public void testSpecifyRowAndColStayA1Write() throws IOException {
        List<ListObjectSheetTest.Item> list = ListObjectSheetTest.Item.randomTestData(10);
        new Workbook().setAutoSize(true)
            .addSheet(new ListSheet<ListObjectSheetTest.Item>("Item"
                , new Column("id").setColIndex(3)
                , new Column("name").setColIndex(4))
                .setData(list)
                .setStartRowIndex(4, false)
            ).writeTo(defaultTestPath.resolve("test specify row and cel stay A1 ListSheet.xlsx"));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("test specify row and cel stay A1 ListSheet.xlsx"))) {
            List<ListObjectSheetTest.Item> readList = reader.sheet(0).bind(ListObjectSheetTest.Item.class, 4).rows().map(row -> (ListObjectSheetTest.Item) row.get()).collect(Collectors.toList());
            assert list.size() == readList.size();
            for (int i = 0, len = list.size(); i < len; i++)
                assert list.get(i).equals(readList.get(i));
        }
    }

    @Test public void testSpecifyRowIgnoreHeaderWrite() throws IOException {
        List<ListObjectSheetTest.Item> list = ListObjectSheetTest.Item.randomTestData();
        new Workbook().setAutoSize(true)
            .addSheet(new ListSheet<>(list).setStartRowIndex(4).ignoreHeader())
            .writeTo(defaultTestPath.resolve("test specify row 4 ignore header ListSheet.xlsx"));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("test specify row 4 ignore header ListSheet.xlsx"))) {
            List<ListObjectSheetTest.Item> readList = reader.sheet(0)
                .header(3)
                .bind(ListObjectSheetTest.Item.class, new HeaderRow().with(createHeaderRow()))
                .rows()
                .map(row -> (ListObjectSheetTest.Item) row.get())
                .collect(Collectors.toList());
            assert list.size() == readList.size();
            for (int i = 0, len = list.size(); i < len; i++)
                assert list.get(i).equals(readList.get(i));
        }
    }

    @Test public void testSpecifyRowStayA1IgnoreHeaderWrite() throws IOException {
        List<ListObjectSheetTest.Item> list = ListObjectSheetTest.Item.randomTestData();
        new Workbook().setAutoSize(true)
            .addSheet(new ListSheet<>(list).setStartRowIndex(4, false).ignoreHeader())
            .writeTo(defaultTestPath.resolve("test specify row 4 stay A1 ignore header ListSheet.xlsx"));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("test specify row 4 stay A1 ignore header ListSheet.xlsx"))) {
            List<ListObjectSheetTest.Item> readList = reader.sheet(0).rows().map(row -> {
                ListObjectSheetTest.Item e = new ListObjectSheetTest.Item();
                e.setId(row.getInt(0));
                e.setName(row.getString(1));
                return e;
            }).collect(Collectors.toList());
            assert list.size() == readList.size();
            for (int i = 0, len = list.size(); i < len; i++)
                assert list.get(i).equals(readList.get(i));
        }
    }

    @Test public void testSpecifyRowAndColIgnoreHeaderWrite() throws IOException {
        List<ListObjectSheetTest.Item> list = ListObjectSheetTest.Item.randomTestData(10);
        new Workbook().setAutoSize(true)
            .addSheet(new ListSheet<ListObjectSheetTest.Item>("Item"
                , new Column("id").setColIndex(3)
                , new Column("name").setColIndex(4))
                .setData(list)
                .setStartRowIndex(4)
                .ignoreHeader()
            ).writeTo(defaultTestPath.resolve("test specify row and cel ignore header ListSheet.xlsx"));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("test specify row and cel ignore header ListSheet.xlsx"))) {
            List<ListObjectSheetTest.Item> readList = reader.sheet(0).rows().map(row -> {
                ListObjectSheetTest.Item e = new ListObjectSheetTest.Item();
                e.setId(row.getInt(3));
                e.setName(row.getString(4));
                return e;
            }).collect(Collectors.toList());
            assert list.size() == readList.size();
            for (int i = 0, len = list.size(); i < len; i++)
                assert list.get(i).equals(readList.get(i));
        }
    }

    @Test public void testSpecifyRowAndColStayA1IgnoreHeaderWrite() throws IOException {
        List<ListObjectSheetTest.Item> list = ListObjectSheetTest.Item.randomTestData(10);
        new Workbook().setAutoSize(true)
            .addSheet(new ListSheet<ListObjectSheetTest.Item>("Item"
                , new Column("id").setColIndex(3)
                , new Column("name").setColIndex(4))
                .setData(list)
                .setStartRowIndex(4, false)
                .ignoreHeader()
            ).writeTo(defaultTestPath.resolve("test specify row and cel stay A1 ignore header ListSheet.xlsx"));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("test specify row and cel stay A1 ignore header ListSheet.xlsx"))) {
            List<ListObjectSheetTest.Item> readList = reader.sheet(0).rows().map(row -> {
                ListObjectSheetTest.Item e = new ListObjectSheetTest.Item();
                e.setId(row.getInt(3));
                e.setName(row.getString(4));
                return e;
            }).collect(Collectors.toList());
            assert list.size() == readList.size();
            for (int i = 0, len = list.size(); i < len; i++)
                assert list.get(i).equals(readList.get(i));
        }
    }

    @Test public void testCustomerRowHeight() throws IOException {
        List<Template> list = new ArrayList<>();
        list.add(Template.of("备注说明\r\n第二行\r\n第三行\r\n第四行", "岗位名称", "岁位"));
        list.add(Template.of("字段名称", "*岗位名称", "岗位描述"));
        list.add(Template.of("示例", "生产统计员", "按照产品规格、价格、工序、员工、车间等不同对象和要求进行统计数据资料分析"));

        new Workbook().addSheet(
            new ListSheet<>(list).setStyleProcessor(new TemplateStyleProcessor())
                .setRowHeight(62.25D)
                .cancelZebraLine().ignoreHeader().putExtProp(Const.ExtendPropertyKey.MERGE_CELLS, Collections.singletonList(Dimension.of("A1:B1")))
        ).writeTo(defaultTestPath.resolve("Customer row height.xlsx"));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("Customer row height.xlsx"))) {
            Iterator<org.ttzero.excel.reader.Row> iter = reader.sheet(0).rows().iterator();
            assert iter.hasNext();
            org.ttzero.excel.reader.Row row0 = iter.next();
            assert list.get(0).equals(Template.of(row0.getString(0), row0.getString(1), row0.getString(2)));
            Styles styles = row0.getStyles();
            int styleIndex = row0.getCellStyle(0);
            Fill fill0 = styles.getFill(styleIndex), fill1 = styles.getFill(row0.getCellStyle(1)), fill2 = styles.getFill(row0.getCellStyle(2));
            assert fill0 != null && fill0.getPatternType() == PatternType.solid && fill0.getFgColor().equals(new Color(188, 219, 162));
            assert fill1 == null || fill1.getPatternType() == PatternType.none;
            assert fill2 == null || fill2.getPatternType() == PatternType.none;

            assert iter.hasNext();
            org.ttzero.excel.reader.Row row1 = iter.next();
            assert list.get(1).equals(Template.of(row1.getString(0), row1.getString(1), row1.getString(2)));
            org.ttzero.excel.entity.style.Font font0 = styles.getFont(row1.getCellStyle(0)), font1 = styles.getFont(row1.getCellStyle(1)), font2 = styles.getFont(row1.getCellStyle(2));
            assert font0.isBold();
            assert font1.isBold();
            assert font2.isBold();
            assert styles.getHorizontal(row1.getCellStyle(0)) == Horizontals.LEFT;
            assert styles.getHorizontal(row1.getCellStyle(1)) == Horizontals.CENTER;
            assert styles.getHorizontal(row1.getCellStyle(2)) == Horizontals.CENTER;

            assert iter.hasNext();
            org.ttzero.excel.reader.Row row2 = iter.next();
            assert list.get(2).equals(Template.of(row2.getString(0), row2.getString(1), row2.getString(2)));
            assert styles.getHorizontal(row2.getCellStyle(0)) == Horizontals.LEFT;
            assert styles.getHorizontal(row2.getCellStyle(1)) == Horizontals.CENTER;
            assert styles.getHorizontal(row2.getCellStyle(2)) == Horizontals.LEFT;
        }
    }

    @Test public void testTileWriter() throws IOException {
        String fileName = "Dynamic title.xlsx";
        List<TileEntity> data = TileEntity.randomTestData();
        new Workbook().cancelZebraLine().addSheet(new ListSheet<>(data).setSheetWriter(new TileXMLWorksheetWriter(3, LocalDate.now().toString()))).writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            Iterator<org.ttzero.excel.reader.Row> iter = reader.sheet(0).rows().iterator();
            assert iter.hasNext();
            assert (LocalDate.now() +  " 拣货单").equals(iter.next().getString(0));

            assert iter.hasNext();
            assert "差异 | 序号 | 商品 | 数量 | 差异 | 序号 | 商品 | 数量 | 差异 | 序号 | 商品 | 数量".equals(iter.next().toString());

            // TODO assert row data
        }
    }

    @Test public void testEmptySheetSubClassSpecified() throws IOException {
        String fileName = "sub-class specified types.xlsx";
        List<ListObjectSheetTest.Item> expectList = new ArrayList<>();
        new Workbook().setAutoSize(true)
            .addSheet(new ListSheet<ListObjectSheetTest.Item>() {
                int i = 0;
                @Override
                protected List<ListObjectSheetTest.Item> more() {
                    List<ListObjectSheetTest.Item> list = i++ < 1 ? ListObjectSheetTest.Item.randomTestData(10) : null;
                    if (list != null) expectList.addAll(list);
                    return list;
                }
            })
            .writeTo(defaultTestPath.resolve(fileName));

        // Check header row
        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            List<ListObjectSheetTest.Item> list = reader.sheet(0).dataRows().map(row -> row.to(ListObjectSheetTest.Item.class)).collect(Collectors.toList());
            assert list.size() == expectList.size();
            for (int i = 0, len = expectList.size(); i < len; i++) {
               ListObjectSheetTest.Item expect = expectList.get(i), e = list.get(i);
               assert expect.equals(e);
            }
        }
    }

    @Test public void testSpecifyActualClass() throws IOException {
        String fileName = "specify unrelated class.xlsx";
        new Workbook()
            .addSheet(new ListSheet<>().setClass(SubModel.class))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            Iterator<org.ttzero.excel.reader.Row> iter = reader.sheet(0).iterator();
            assert iter.hasNext();
            org.ttzero.excel.reader.Row row = iter.next();
            assert "name".equals(row.getString(0));
            assert "score".equals(row.getString(1));
        }
    }

    @Test public void testSpecifyConvertClass() throws IOException {
        List<SpecifyConvertModel> expectList = SpecifyConvertModel.randomTestData(20);
        String fileName = "specify converter test.xlsx";
        new Workbook()
            .addSheet(new ListSheet<>(expectList))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            List<SpecifyConvertModel> readList = reader.sheet(0).header(1).rows().map(row -> row.to(SpecifyConvertModel.class)).collect(Collectors.toList());
            assert expectList.size() == readList.size();
            for (int i = 0, len = expectList.size(); i < len; i++)
                assert expectList.get(i).equals(readList.get(i));
        }
    }

    @Test public void testAutoSize() throws IOException {
        String fileName = "test auto size.xlsx";
        List<ListObjectSheetTest.Student> expectList = ListObjectSheetTest.Student.randomTestData();
        new Workbook()
            .setAutoSize(true)
            .addSheet(new ListSheet<>(expectList
                , new Column("学号", "id").setStyleProcessor((o, style, sst)
                    -> (((int) o & 1) == 1 ? sst.modifyFont(style, new Font("Algerian", 24)) : ((int) o) < 10 ? sst.modifyFont(style, new Font("Algerian", 56)) : style))
                , new Column("姓名", "name").setStyleProcessor((o, style, sst) -> {
                    int len = ((String) o).length();
                    if (len < 5) {
                        style = sst.modifyFont(style, new Font("Trebuchet MS", 72));
                    } else if (len > 15) {
                        style = sst.modifyFont(style, new Font("宋体", 5));
                    } else if (len > 10) {
                        style = sst.modifyFont(style, new Font("Bauhaus 93", 18));
                    }
                    return style;
                })
            ))
            .writeTo(defaultTestPath.resolve(fileName));
    }

    @Test public void testCustomStyle() throws IOException {
        String fileName = "test custom style.xlsx";
        List<ListObjectSheetTest.Student> expectList = ListObjectSheetTest.Student.randomTestData();
        new Workbook()
            .addSheet(new ListSheet<>(expectList
                , new Column("id").setFont(new Font("微软雅黑", 16)).setHorizontal(Horizontals.CENTER)
                , new Column("name").setFont(new Font("华文行楷", 23)).setBorder(new Border()).autoSize()
            ))
            .writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            Sheet sheet = reader.sheet(0);
            List<ListObjectSheetTest.Student> list = sheet.forceImport().dataRows().map(row -> row.to(ListObjectSheetTest.Student.class)).collect(Collectors.toList());
            assert list.size() == expectList.size();
            for (int i = 0; i < expectList.size(); i++) {
                ListObjectSheetTest.Student e = expectList.get(i), o = list.get(i);
                assert e.getName().equals(o.getName());
                assert e.getId() == o.getId();
            }

            for (Iterator<org.ttzero.excel.reader.Row> iter = sheet.reset().dataRows().iterator(); iter.hasNext(); ) {
                org.ttzero.excel.reader.Row row = iter.next();
                Styles styles = row.getStyles();
                // 第一列样式
                {
                    int style = row.getCellStyle(0);
                    Font font = styles.getFont(style);
                    assert "微软雅黑".equals(font.getName());
                    assert font.getSize() == 16;
                    int horizontal = styles.getHorizontal(style);
                    assert horizontal == Horizontals.CENTER;
                }
                // 第二列样式
                {
                    int style = row.getCellStyle(1);
                    Font font = styles.getFont(style);
                    assert "华文行楷".equals(font.getName());
                    assert font.getSize() == 23;
                    int horizontal = styles.getHorizontal(style);
                    assert horizontal == Horizontals.LEFT;
                    Border border = styles.getBorder(style);
                    assert border == null || border.getBorderTop().getStyle() == BorderStyle.NONE;
                }
            }
        }
    }

    public static class TemplateStyleProcessor implements StyleProcessor<Template> {
        String k;
        int c = 0;
        @Override
        public int build(Template o, int style, Styles sst) {
            if (!o.v1.equals(k)) {
                k = o.v1;
                c = 0;
            }
            if (o.v1.startsWith("备注说明")) {
                if (c == 0)
                    style = sst.modifyFill(style, new Fill(PatternType.solid, new Color(188, 219, 162)));
            }
            else if (o.v1.equals("字段名称")) {
                Font font = sst.getFont(style);
                style = sst.modifyFont(style, font.clone().bold());
                if (c > 0)
                    style = sst.modifyHorizontal(style, Horizontals.CENTER);
            }
            else if (o.v1.equals("示例")) {
                if (c == 1)
                    style = sst.modifyHorizontal(style, Horizontals.CENTER);
            }
            c++;
            return style;
        }
    }

    public static class TileEntity {
        @ExcelColumn("{date} 拣货单")
        @ExcelColumn(value = "差异", maxWidth = 8.6D)
        private String diff;
        @ExcelColumn("{date} 拣货单")
        @ExcelColumn(value = "序号", maxWidth = 6.8D)
        private Integer no;
        @ExcelColumn("{date} 拣货单")
        @ExcelColumn(value = "商品", maxWidth = 12.0D)
        private String product;
        @ExcelColumn("{date} 拣货单")
        @ExcelColumn(value = "数量", maxWidth = 6.8D)
        private Integer num;

        public static List<TileEntity> randomTestData() {
            int n = 23;
            List<TileEntity> list = new ArrayList<>(n);
            for (int i = 0; i < n; i++) {
                TileEntity e = new TileEntity();
                e.no = i + 1;
                e.product = getRandomString(10);
                e.num = random.nextInt(20) + 1;
                list.add(e);
            }
            return list;
        }
    }

    /**
     * 自定义平铺WorksheetWriter
     */
    public static class TileXMLWorksheetWriter extends XMLWorksheetWriter {
        private int tile; // 平铺的数量，也就是每行重复输出多少条数据
        private String date; // 可忽略，仅仅是表头上的日期

        public TileXMLWorksheetWriter(int tile) {
            this.tile = tile;
        }

        public TileXMLWorksheetWriter(int tile, String date) {
            this.tile = tile;
            this.date = date;
        }

        public int getTile() {
            return tile;
        }

        public void setTile(int tile) {
            this.tile = tile;
        }

        public String getDate() {
            return date;
        }

        public void setDate(String date) {
            this.date = date;
        }

        @Override
        protected void writeBefore() throws IOException {
            // The header columns
            columns = sheet.getAndSortHeaderColumns();
            // Give new columns
            tileColumns();

            boolean nonHeader = sheet.getNonHeader() == 1;

            bw.write(Const.EXCEL_XML_DECLARATION);
            // Declaration
            bw.newLine();
            // Root node
            writeRootNode();

            // Dimension
            writeDimension();

            // SheetViews default value
            writeSheetViews();

            // Default row height and width
            int fillSpace = 6;
            BigDecimal width = BigDecimal.valueOf(!nonHeader ? sheet.getDefaultWidth() : 8.38D);
            String defaultWidth = width.setScale(2, BigDecimal.ROUND_HALF_UP).toString();
            writeSheetFormat(fillSpace, defaultWidth);

            // cols
            writeCols(fillSpace, defaultWidth);
        }

        protected void tileColumns() {
            if (tile == 1) return;

            int x = columns.length, y = x * tile, t = columns[columns.length - 1].getRealColIndex();
            // Bound check
            if (y > Const.Limit.MAX_COLUMNS_ON_SHEET)
                throw new TooManyColumnsException(y, Const.Limit.MAX_COLUMNS_ON_SHEET);

            Column[] _columns = new Column[y];
            for (int i = 0; i < y; i++) {
                // 第一个对象的表头不需要复制
                Column col = i < x ? columns[i] : new Column(columns[i % x]).addSubColumn(new Column());
                col.realColIndex = columns[i % x].realColIndex + t * (i / x);
                _columns[i] = col;

                // 替换拣货单上的日期
                Column _col = col;
                do {
                    if (StringUtil.isNotEmpty(_col.getName()) && _col.getName().contains("{date}"))
                        _col.setName(_col.getName().replace("{date}", date));
                }
                while ((_col = _col.next) != null);
            }

            this.columns = _columns;

            // FIXME 这里强行指定合并替换掉原本的头
            List<Dimension> mergeCells = Collections.singletonList(new Dimension(1, (short) 1, 1, (short) y));
            sheet.putExtProp(Const.ExtendPropertyKey.MERGE_CELLS, mergeCells);
        }

        @Override
        protected void writeRow(Row row) throws IOException {
            Cell[] cells = row.getCells();
            int len = cells.length, r = row.getIndex() / tile + startRow, c = columns[columns.length - 1].realColIndex / tile, y = row.getIndex() % tile;
            if (y == 0) startRow(r - startRow, columns[columns.length - 1].realColIndex, -1D);

            for (int i = 0; i < len; i++) {
                Cell cell = cells[i];
                int xf = cell.xf, col = i + c * y;
                switch (cell.t) {
                    case INLINESTR:
                    case SST:
                        writeString(cell.sv, r, col, xf);
                        break;
                    case NUMERIC:
                        writeNumeric(cell.nv, r, col, xf);
                        break;
                    case LONG:
                        writeNumeric(cell.lv, r, col, xf);
                        break;
                    case DATE:
                    case DATETIME:
                    case DOUBLE:
                    case TIME:
                        writeDouble(cell.dv, r, col, xf);
                        break;
                    case BOOL:
                        writeBool(cell.bv, r, col, xf);
                        break;
                    case DECIMAL:
                        writeDecimal(cell.mv, r, col, xf);
                        break;
                    case CHARACTER:
                        writeChar(cell.cv, r, col, xf);
                        break;
                    case BLANK:
                        writeNull(r, col, xf);
                        break;
                    default:
                }
            }
            // 注意这里可能不会关闭row需要在writeAfter进行二次处理
            if (y == tile - 1)
                bw.write("</row>");
        }

        @Override
        protected void writeAfter(int total) throws IOException {
            if (total > 0 && (total - 1) % tile < tile - 1) bw.write("</row>");
            super.writeAfter(total);
        }
    }


    private static org.ttzero.excel.reader.Row createHeaderRow () {
        org.ttzero.excel.reader.Row headerRow = new org.ttzero.excel.reader.Row() {};
        Cell[] cells = new Cell[2];
        cells[0] = new Cell((short) 1).setSv("id");
        cells[1] = new Cell((short) 2).setSv("name");
        headerRow.setCells(cells);
        return headerRow;
    }


    public static class Template {
        @ExcelColumn(maxWidth = 12.0D, wrapText = true)
        String v1;
        @ExcelColumn(maxWidth = 20.0, wrapText = true)
        String v2;
        @ExcelColumn(maxWidth = 25.0D, wrapText = true)
        String v3;

        static Template of(String v1, String v2, String v3) {
            Template v = new Template();
            v.v1 = v1;
            v.v2 = v2;
            v.v3 = v3;
            return v;
        }

        @Override
        public boolean equals(Object o) {
            if (this == o) return true;
            if (o == null || getClass() != o.getClass()) return false;
            Template template = (Template) o;
            return Objects.equals(v1, template.v1) &&
                Objects.equals(v2, template.v2) &&
                Objects.equals(v3, template.v3);
        }

        @Override
        public int hashCode() {
            return Objects.hash(v1, v2, v3);
        }
    }

    public static class SubModel {
        @ExcelColumn
        private String name;
        @ExcelColumn
        private int status;
    }

    public static class SpecifyConvertModel {
        @ExcelColumn
        private String name;
        @ExcelColumn(converter = StatusConvert.class)
        private int status;

        public static List<SpecifyConvertModel> randomTestData(int n) {
            List<SpecifyConvertModel> list = new ArrayList<>(n);
            for (int i = 0; i < n; i++) {
                SpecifyConvertModel e = new SpecifyConvertModel();
                e.name = getRandomString(10);
                e.status = random.nextInt(4);
                list.add(e);
            }
            return list;
        }

        @Override
        public boolean equals(Object o) {
            if (this == o) return true;
            if (o == null || getClass() != o.getClass()) return false;
            SpecifyConvertModel that = (SpecifyConvertModel) o;
            return status == that.status && Objects.equals(name, that.name);
        }

        @Override
        public int hashCode() {
            return Objects.hash(name, status);
        }
    }

    public static class StatusConvert implements Converter<Integer> {
        final String[] statusDesc = { "未开始", "进行中", "完结", "中止" };

        @Override
        public Integer reversion(String v) {
            for (int i = 0; i < statusDesc.length; i++) {
                if (statusDesc[i].equals(v)) {
                    return i;
                }
            }
            return null;
        }

        @Override
        public Object conversion(Object v) {
            return v != null ? statusDesc[(int) v] : null;
        }
    }
}

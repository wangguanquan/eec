package org.ttzero.excel.annotation;

import org.junit.Test;
import org.ttzero.excel.entity.ListSheet;
import org.ttzero.excel.entity.Workbook;
import org.ttzero.excel.entity.WorkbookTest;
import org.ttzero.excel.entity.style.Fill;
import org.ttzero.excel.entity.style.PatternType;
import org.ttzero.excel.entity.style.Styles;
import org.ttzero.excel.reader.Cell;
import org.ttzero.excel.reader.ExcelReader;
import org.ttzero.excel.reader.Row;

import java.io.IOException;
import java.util.Collections;
import java.util.Iterator;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;

/**
 * 自定义导出Excel表头样式
 *
 * @author carljia
 */
public class HeaderStyleTest extends WorkbookTest {

    @Test
    public void testOriginal() throws IOException {
        String fileName = "customize_header_style_original.xlsx";
        Head itemFull = new Head();
        new Workbook().setAutoSize(true).addSheet(new ListSheet<>(Collections.singletonList(itemFull))).writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            Styles styles = reader.getStyles();
            Iterator<Row> iter = reader.sheet(0).iterator();
            assertTrue(iter.hasNext());
            Row row = iter.next();
            String[] titles = {"column1", "column2", "column3", "code", "错误信息"};
            for (int i = row.getFirstColumnIndex(); i < row.getLastColumnIndex(); i++) {
                Cell cell = row.getCell(i);
                assertEquals(row.getString(cell), titles[i]);
                int style = row.getCellStyle(cell);
                Fill fill = styles.getFill(style);
                assertEquals(fill.getPatternType(), PatternType.solid);
                assertEquals(fill.getFgColor(), Styles.toColor("#E9EAEC"));
            }
        }
    }

    @Test
    public void testFillBgColor() throws IOException {
        String fileName = "customize_header_style_bgc.xlsx";
        Head1 itemFull = new Head1();
        new Workbook().setAutoSize(true).addSheet(new ListSheet<>(Collections.singletonList(itemFull))).writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            Styles styles = reader.getStyles();
            Iterator<Row> iter = reader.sheet(0).iterator();
            assertTrue(iter.hasNext());
            Row row = iter.next();
            String[] titles = {"column4", "column1", "column2", "column3", "code", "错误信息"};
            for (int i = row.getFirstColumnIndex(); i < row.getLastColumnIndex(); i++) {
                Cell cell = row.getCell(i);
                assertEquals(row.getString(cell), titles[i]);
                int style = row.getCellStyle(cell);
                Fill fill = styles.getFill(style);
                assertEquals(fill.getPatternType(), PatternType.solid);
                assertEquals(fill.getFgColor(), Styles.toColor(i < row.getLastColumnIndex() - 1 ? "#E9EAEC" : "#ff0000"));
            }
        }
    }

    @Test
    public void testFontColor() throws IOException {
        String fileName = "customize_header_style_fc.xlsx";
        Head2 itemFull = new Head2();
        new Workbook().setAutoSize(true).addSheet(new ListSheet<>(Collections.singletonList(itemFull))).writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            Styles styles = reader.getStyles();
            Iterator<Row> iter = reader.sheet(0).iterator();
            assertTrue(iter.hasNext());
            Row row = iter.next();
            String[] titles = {"column4", "column5", "column1", "column2", "column3", "code", "错误信息"};
            for (int i = row.getFirstColumnIndex(); i < row.getLastColumnIndex(); i++) {
                Cell cell = row.getCell(i);
                assertEquals(row.getString(cell), titles[i]);
                int style = row.getCellStyle(cell);
                Fill fill = styles.getFill(style);
                assertEquals(fill.getPatternType(), PatternType.solid);
                assertEquals(fill.getFgColor(), Styles.toColor(i != 1 ? "#E9EAEC" : "#cccccc"));

                if (i == row.getLastColumnIndex() - 1)
                    assertEquals(styles.getFont(style).getColor(), Styles.toColor("#ff0000"));
            }
        }
    }

    @Test public void testAnnoOnClassTest() throws IOException {
        String fileName = "annotation on class.xlsx";
        Head3 itemFull = new Head3();
        new Workbook().setAutoSize(true).addSheet(new ListSheet<>(Collections.singletonList(itemFull))).writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            Styles styles = reader.getStyles();
            Iterator<Row> iter = reader.sheet(0).iterator();
            assertTrue(iter.hasNext());
            Row row = iter.next();
            String[] titles = {"column1", "column2", "column3", "code", "错误信息"};
            for (int i = row.getFirstColumnIndex(); i < row.getLastColumnIndex(); i++) {
                Cell cell = row.getCell(i);
                assertEquals(row.getString(cell), titles[i]);
                int style = row.getCellStyle(cell);
                Fill fill = styles.getFill(style);
                assertEquals(fill.getPatternType(), PatternType.solid);
                assertEquals(fill.getFgColor(), Styles.toColor("#ffff00"));
                assertEquals(styles.getFont(style).getColor(), Styles.toColor("red"));
            }
        }
    }

    @Test public void testAnnoOnClassAndMethodTest() throws IOException {
        String fileName = "annotation on class and method.xlsx";
        Head4 itemFull = new Head4();
        new Workbook().setAutoSize(true).addSheet(new ListSheet<>(Collections.singletonList(itemFull))).writeTo(defaultTestPath.resolve(fileName));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(fileName))) {
            Styles styles = reader.getStyles();
            Iterator<Row> iter = reader.sheet(0).iterator();
            assertTrue(iter.hasNext());
            Row row = iter.next();
            String[] titles = {"column1", "column2", "column3", "code", "错误信息"};
            for (int i = row.getFirstColumnIndex(); i < row.getLastColumnIndex(); i++) {
                Cell cell = row.getCell(i);
                assertEquals(row.getString(cell), titles[i]);
                int style = row.getCellStyle(cell);
                Fill fill = styles.getFill(style);
                assertEquals(fill.getPatternType(), PatternType.solid);
                if (i < row.getLastColumnIndex() - 1) {
                    assertEquals(fill.getFgColor(), Styles.toColor("#ffff00"));
                } else {
                    assertEquals(styles.getFont(style).getColor(), Styles.toColor("blue"));
                    assertEquals(fill.getFgColor(), Styles.toColor("#E9EAEC"));
                }
            }
        }
    }

    private static class Head {
        @ExcelColumn
        private String column1;
        @ExcelColumn
        private String column2;
        @ExcelColumn
        private String column3;
        @ExcelColumn
        private String code;
        /**
         * errorMsg
         */
        @ExcelColumn(value = "错误信息")
        private String errorMsg;

        public String getColumn1() {
            return column1;
        }

        public String getColumn2() {
            return column2;
        }

        public String getColumn3() {
            return column3;
        }

        public String getCode() {
            return code;
        }

        public String getErrorMsg() {
            return errorMsg;
        }
    }

    private static class Head1 extends Head {
        @ExcelColumn
        private String column4;
        /**
         * errorMsg
         */
        @ExcelColumn(value = "错误信息")
        @HeaderStyle(fillFgColor = "#ff0000")
        public String getErrorMsg() {
            return null;
        }
    }

    private static class Head2 extends Head {
        @ExcelColumn(value = "错误信息")
        @HeaderStyle(fontColor = "#ff0000")
        public String getErrorMsg() {
            return null;
        }

        @ExcelColumn
        @HeaderStyle(fontColor = "black")
        private String column4;
        @ExcelColumn
        @HeaderStyle(fillFgColor = "#cccccc")
        private String column5;
    }

    @HeaderStyle(fontColor = "red", fillFgColor = "#ffff00")
    private static class Head3 extends Head { }

    @HeaderStyle(fontColor = "red", fillFgColor = "#ffff00")
    private static class Head4 extends Head {
        @Override
        @HeaderStyle(fontColor = "blue")
        public String getErrorMsg() {
            return super.getErrorMsg();
        }
    }
}

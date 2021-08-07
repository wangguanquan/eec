package org.ttzero.excel.annotation;

import org.junit.Test;
import org.ttzero.excel.entity.ListSheet;
import org.ttzero.excel.entity.Workbook;
import org.ttzero.excel.entity.WorkbookTest;

import java.io.IOException;
import java.util.Collections;

/**
 * 自定义导出Excel表头样式
 */
public class HeaderStyleTest extends WorkbookTest {

    @Test
    public void testOriginal() throws IOException {
        Head itemFull = new Head();
        new Workbook("customize_header_style_original").setAutoSize(true).addSheet(new ListSheet<>(Collections.singletonList(itemFull))).writeTo(defaultTestPath);

    }

    @Test
    public void testFillBgColor() throws IOException {
        Head1 itemFull = new Head1();
        new Workbook("customize_header_style_bgc").setAutoSize(true).addSheet(new ListSheet<>(Collections.singletonList(itemFull))).writeTo(defaultTestPath);

    }

    @Test
    public void testFontColor() throws IOException {
        Head2 itemFull = new Head2();
        new Workbook("customize_header_style_fc").setAutoSize(true).addSheet(new ListSheet<>(Collections.singletonList(itemFull))).writeTo(defaultTestPath);

    }

    @Test public void testAnnoOnClassTest() throws IOException {
        Head3 itemFull = new Head3();
        new Workbook("annotation on class").setAutoSize(true).addSheet(new ListSheet<>(Collections.singletonList(itemFull))).writeTo(defaultTestPath);
    }

    @Test public void testAnnoOnClassAndMethodTest() throws IOException {
        Head4 itemFull = new Head4();
        new Workbook("annotation on class and method").setAutoSize(true).addSheet(new ListSheet<>(Collections.singletonList(itemFull))).writeTo(defaultTestPath);
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

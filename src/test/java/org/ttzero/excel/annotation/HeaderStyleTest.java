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
    public void originalTest() throws IOException {
        Head itemFull = new Head();
        new Workbook("customize_header_style_original").setAutoSize(true).addSheet(new ListSheet<>(Collections.singletonList(itemFull))).writeTo(defaultTestPath);

    }

    @Test
    public void fillBgColorTest() throws IOException {
        Head1 itemFull = new Head1();
        new Workbook("customize_header_style_bgc").setAutoSize(true).addSheet(new ListSheet<>(Collections.singletonList(itemFull))).writeTo(defaultTestPath);

    }

    @Test
    public void fontColorTest() throws IOException {
        Head2 itemFull = new Head2();
        new Workbook("customize_header_style_fc").setAutoSize(true).addSheet(new ListSheet<>(Collections.singletonList(itemFull))).writeTo(defaultTestPath);

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
}

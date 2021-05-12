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

    static class Head {

        /**
         * errorMsg
         */
        @ExcelColumn(value = "错误信息")
        private String errorMsg;

    }

    static class Head1 {

        /**
         * errorMsg
         */
        @ExcelColumn(value = "错误信息")
        @HeaderStyle(fillBgColor = "#ff0000")
        private String errorMsg;

    }

    static class Head2 {

        /**
         * errorMsg
         */
        @ExcelColumn(value = "错误信息")
        @HeaderStyle(fontColor = "#ff0000")
        private String errorMsg;

    }
}

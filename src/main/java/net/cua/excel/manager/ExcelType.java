package net.cua.excel.manager;

/**
 * Type of excel. Biff8 or XLSX(xml zip)
 * Create by guanquan.wang at 2019-01-24 10:12
 */
public enum ExcelType {
    /**
     * Excel 8.0
     * Excel 9.0
     * Excel 10.0
     * Excel 11.0
     */
    BIFF8,
    /**
     * Excel 12.0~
     */
    XLSX,
    /**
     * Others
     */
    UNKNOWN
}

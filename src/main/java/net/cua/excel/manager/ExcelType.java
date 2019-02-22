package net.cua.excel.manager;

/**
 * Type of excel. Biff8 or XLSX(Office open xml)
 * Create by guanquan.wang at 2019-01-24 10:12
 */
public enum ExcelType {
    /**
     * BIFF8 only
     * Excel 8.0 Excel 97
     * Excel 9.0 Excel 2000
     * Excel 10.0 Excel XP
     * Excel 11.0 Excel 2003
     */
    XLS,

    /**
     * Excel 12.0~
     */
    XLSX,

    /**
     * Others
     */
    UNKNOWN
}

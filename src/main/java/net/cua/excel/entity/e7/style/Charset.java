package net.cua.excel.entity.e7.style;

/**
 * Created by guanquan.wang at 2018-02-08 17:05
 */
public class Charset {
    public static final int ANSI = 0x00
            , DEFAULT = 0x01  // not a real charset; rather
            // , it is a constant akin to NULL that means
            // "show characters in whatever charsets are available."
            , SYMBOL = 0x02
            , SHIFTJIS = 0x80 // 日文
            , HANGUL = 0x81 // 韩国、朝鲜
            , GB2312 = 0x86 // 简体
            , CHINESEBIG5 = 0x88 // 繁体
            , GREEK = 0xA1
            , TURKISH = 0xA2
            , HEBREW = 0xB1
            , ARABIC = 0xB2
            , BALTIC = 0xBA
            , RUSSIAN = 0xCC
            , THAI = 0xDE
            , EE = 0xEE
            , OEM = 0xFF
            ;
}

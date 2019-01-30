package net.cua.excel.entity.e7.style;

/**
 * Created by guanquan.wang at 2018-02-08 17:05
 */
public class Charset {
    public static final int ANSI = 0x00 // ANSI Latin
            , DEFAULT = 0x01  // not a real charset; rather
            // , it is a constant akin to NULL that means
            // "show characters in whatever charsets are available."
            , SYMBOL = 0x02 //  Symbol
            , SHIFTJIS = 0x80 // 日文
            , HANGUL = 0x81 // 韩国、朝鲜
			, JOHAB = 0x82 //  ANSI Korean (Johab)
            , GB2312 = 0x86 // 简体
            , CHINESEBIG5 = 0x88 // 繁体
            , GREEK = 0xA1 //  ANSI Greek
            , TURKISH = 0xA2 // ANSI Turkish
            , HEBREW = 0xB1 // ANSI Hebrew
            , ARABIC = 0xB2 // ANSI Arabic
            , BALTIC = 0xBA // ANSI Baltic
            , RUSSIAN = 0xCC // ANSI Cyrillic
            , THAI = 0xDE //  ANSI Thai
            , EE = 0xEE // ANSI Latin II (Central European)
            , OEM = 0xFF // OEM Latin I
            ;
}

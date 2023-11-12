/*
 * Copyright (c) 2017-2018, guanquan.wang@yandex.com All Rights Reserved.
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

package org.ttzero.excel.entity.style;

/**
 * 字符集
 *
 * @author guanquan.wang at 2018-02-08 17:05
 */
public class Charset {
    /**
     * 0 标准 Windows 字符 (ANSI)
     */
    public static final int ANSI = 0x00;
    /**
     * 1 默认字符集
     * 不是真正的字符集；它是一个类似于NULL的常量，意思是“显示可用字符集中的字符”
     */
    public static final int DEFAULT = 0x01;
    /**
     * 2 符号字符集
     */
    public static final int SYMBOL = 0x02;
    /**
     * 128 日文字符集
     */
    public static final int SHIFTJIS = 0x80;
    /**
     * 129 朝鲜语字符集(韩国、朝鲜)
     */
    public static final int HANGUL = 0x81;
    /**
     * 130 朝鲜语字符集(Johab)
     */
    public static final int JOHAB = 0x82;
    /**
     * 134 简体中文字符集
     */
    public static final int GB2312 = 0x86;
    /**
     * 136 繁体中文字符集
     */
    public static final int CHINESEBIG5 = 0x88;
    /**
     * 161 希腊字符集
     */
    public static final int GREEK = 0xA1;
    /**
     * 162 土耳其字符集
     */
    public static final int TURKISH = 0xA2;
    /**
     * 177 希伯来语字符集
     */
    public static final int HEBREW = 0xB1;
    /**
     * 178 阿拉伯语字符集
     */
    public static final int ARABIC = 0xB2;
    /**
     * 186 印欧语系中的）波罗的语族
     */
    public static final int BALTIC = 0xBA;
    /**
     * 204 西里尔文 用于俄语、保加利亚语及其他一些中欧语言
     */
    public static final int RUSSIAN = 0xCC;
    /**
     * 222 泰国字符集
     */
    public static final int THAI = 0xDE;
    /**
     * 238 拉丁语II（中欧）
     */
    public static final int EE = 0xEE;
    /**
     * 255 通常由 Microsoft MS-DOS 应用程序显示的扩展字符
     */
    public static final int OEM = 0xFF;
}

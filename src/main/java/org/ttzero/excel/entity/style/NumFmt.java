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

import org.dom4j.Element;
import org.ttzero.excel.util.StringUtil;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.Comparator;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;

import static org.ttzero.excel.entity.style.Styles.getAttr;

/**
 * To create a custom number format, you start by selecting one of the built-in number formats as a starting point.
 * You can then change any one of the code sections of that format to create your own custom number format.
 * <p>
 * A number format can have up to four sections of code, separated by semicolons.
 * These code sections define the format for positive numbers, negative numbers, zero values, and text, in that order.
 * <p>
 * &lt;POSITIVE&gt;;&lt;NEGATIVE&gt;;&lt;ZERO&gt;;&lt;TEXT&gt;
 * <p>
 * For example, you can use these code sections to create the following custom format:
 * <p>
 * [Blue]#,##0.00_);[Red](#,##0.00);0.00;"sales "@
 * <p>
 * You do not have to include all code sections in your custom number format.
 * If you specify only two code sections for your custom number format,
 * the first section is used for positive numbers and zeros, and the second section is used for negative numbers.
 * If you specify only one code section, it is used for all numbers.
 * If you want to skip a code section and include a code section that follows it,
 * you must include the ending semicolon for the section that you skip.
 * <ul>
 * <li><a href="https://support.office.com/en-us/article/create-or-delete-a-custom-number-format-78f2a361-936b-4c03-8772-09fab54be7f4">Create a custom number format</a></li>
 * <li><a href="https://support.office.com/en-us/article/Number-format-codes-5026bbd6-04bc-48cd-bf33-80f18b4eae68?ui=en-US&rs=en-US&ad=US">Number format codes</a></li>
 * <li><a href="https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2010/ee857658(v=office.14)">NumberingFormat Class</a></li>
 * <li><a href="https://support.microsoft.com/zh-cn/office/%E6%95%B0%E5%AD%97%E6%A0%BC%E5%BC%8F%E4%BB%A3%E7%A0%81-5026bbd6-04bc-48cd-bf33-80f18b4eae68">数字格式代码</a></li>
 * </ul>
 *
 * @author guanquan.wang at 2018-02-06 08:51
 */
public class NumFmt implements Comparable<NumFmt> {

    /**
     * Format as {@code yyyy-mm-dd hh:mm:ss}
     */
    public static final NumFmt DATETIME_FORMAT = new NumFmt("yyyy\\-mm\\-dd\\ hh:mm:ss"),
    /**
     * Format as {@code yyyy-mm-dd}
     */
    DATE_FORMAT = new NumFmt("yyyy\\-mm\\-dd"),
    /**
     * Format as {@code hh:mm:ss}
     */
    TIME_FORMAT = new NumFmt("hh:mm:ss");

    protected String code;
    protected int id = -1;

    public NumFmt() { }

    NumFmt(int id, String code) {
        this.id = id;
        this.code = code;
    }

    public NumFmt(String code) {
        this.code = clean(code);
    }

    public String getCode() {
        return code;
    }

    public void setCode(String code) {
        this.code = code;
    }

    NumFmt setId(int id) {
        this.id = id;
        return this;
    }

    public int getId() {
        return id;
    }

    /**
     * Built-In number format
     *
     * @param id the built-in id
     * @return the {@link NumFmt}
     */
    public static NumFmt valueOf(int id) {
        return new NumFmt().setId(id);
    }

    /**
     * Create a NumFmt
     *
     * @param code the numFmt code string
     * @return NumFmt
     */
    public static NumFmt of(String code) {
        return new NumFmt(code);
    }

    // Clean the format code
    private static String clean(String code) {
        if (StringUtil.isEmpty(code))
            throw new NumberFormatException("The format code must not be null or empty.");

        // Replace '-' to '\-'
        code = escape(code, '-');
        // Replace ' ' to '\ '
        code = escape(code, ' ');

        return code;
    }

    private static String escape(String code, char c) {
        int i = code.indexOf(c);
        if (i > -1) {
            int j = 0;
            StringBuilder buf = new StringBuilder();
            do {
                if (i != j) {
                    buf.append(code, j, i);
                    j = i;
                }
                if (i == 0 || code.charAt(i - 1) != '\\') {
                    buf.append('\\');
                }
            } while ((i = code.indexOf(c, i + 1)) > -1);
            code = buf.append(code, j, code.length()).toString();
        }
        return code;
    }

    /**
     * 兼容之前版本，这里固定按默认字体 “宋体” 11字号处理，后续将删除
     */
    static final Font SONG = new Font("宋体", 11);
    /**
     * 粗略计算单元格长度
     *
     * @param base the cell value length
     * @return cell length
     * @deprecated 使用 {@link #calcNumWidth(double, Font)}替代，新方法会根据字体/字号进行计算
     */
    @Deprecated
    public double calcNumWidth(double base) {
        return calcNumWidth(base, SONG);
    }

    /**
     * 缓存code的宽度
     *
     * key: 字号+字体
     * value: 预计算的结果
     */
    protected transient Map<String, Integer> codeWidthCache;

    /**
     * 粗略计算单元格长度，优先从缓存中获取预处理结果，缓存key由字号+字体名组成这样就保存能计算出相近的宽度，
     * 未命中缓存则从先预处理再丢入缓存以便下次使用
     *
     * @param base the cell value length
     * @param font font
     * @return cell length
     */
    public double calcNumWidth(double base, Font font) {
        if (StringUtil.isBlank(code)) return 0.0D;
        // 获取code预处理后的中间结果
        int widthCache = getCodeWidthFromCache(font);

        double width = 0D;
        // 日期
        if ((widthCache & 1) == 1) width = (widthCache >> 2) / 10000.0;
        else if (base >= 1) {
            int comma = (widthCache >>> 1) & 1, k = (widthCache >>> 2) & 63;
            double s = (widthCache >>> 8) / 10000.0;
            width = (base + (comma == 1 ? (base - 1) / 3 : 1) + k) * s; // 有逗号分隔符时计算分隔符个数
        }
        return width;
    }

    /**
     * 计算并缓存格式化串的长度，以此长度为基础计算文本长度
     *
     * @param font 字体
     * @return 一个二进制结果，第0位表示日期，第1位表示是否有逗号分隔符，当第0位为1时高30位保存格式化串的宽度（已按字体计算好的宽度），
     * 当第0位为0时第2-9位表示小数点后面的位数，10-31位表示单字节单个字符宽度
     */
    protected int getCodeWidthFromCache(Font font) {
        if (codeWidthCache == null) codeWidthCache = new HashMap<>();
        return codeWidthCache.computeIfAbsent(font.getSize() + font.getName(), key -> {
            int wc = 0;
            boolean isDate = Styles.testCodeIsDate(code);
            // 计算每一段的宽度取最大值
            String[] codes = code.split(";");
            int[] ks = new int[codes.length];
            java.awt.FontMetrics fm = font.getFontMetrics();
            /*
             粗略估算单／双字节宽度，与实际计算出来的结果可能有很大区别，输出到Excel的宽度需要除{@code 6}，
             中文的宽度相对简单几乎都是一样的宽度，英文却很复杂较窄的有{@code 'i','l',':'}和部分符号而像
             {@code 'X','E','G'，’%'，‘@’}这类又比较宽，本方法取20个字符平均宽度为单字节宽度，format大多数是数字或数字相关的符号
             所以这里只计算数字和数字相关符号的平均宽度
             */
            double s = fm.stringWidth("1234567890.,: %*-+<>") / 120.0D, d = font.getSize2() / 6.0D;
            for (int i = 0; i < codes.length; i++) {
                String code = codes[i];
                double n = 0.0D;
                boolean ignore = false, comma = false;
                int len = code.length();
                for (int j = 0; j < len; j++) {
                    char c = code.charAt(j);
                    if (c == '"' || c == '\\') continue;
                    if (ignore) {
                        if (c == ']' || c == ')') {
                            ignore = false;
                        }
                        continue;
                    }
                    if (c == '[' || c == '(') {
                        ignore = true;
                        continue;
                    }
                    if (c == ',') comma = true;
                    // 需要使用"方言"为了简单这里只处理am/pm 或者 上午/下午 特殊处理，最终显示只显示其中一个
                    else if (c == '/' && j >= 2 && j + 2 < len) {
                        char p1 = code.charAt(j - 2), p2 = code.charAt(j - 1)
                            , n1 = code.charAt(j + 1), n2 = code.charAt(j + 2);
                        if (p1 == '上' && n1 == '下' && p2 == '午' && n2 == '午'
                            || (p1 == 'a' || p1 == 'A') && (n1 == 'p' || n1 == 'P') && (p2 == 'm' || p2 == 'M') && (n2 == 'm' || n2 == 'M')) {
                            j += 2;
                            continue;
                        }
                    }
                    n += c > 0x4E00 ? d : s;
                }

                // 日期格式，只有一个段
                if (isDate) {
                    wc = ((int) (n * 10000 + 0.5)) << 2;
                    wc |= comma ? 3 : 1;
                    break;
                }
                // 数字格式，可能包含多个段，这里要计算出最长的那个段并进行缓存
                // 整数部分可能添加逗号等分隔符
                else {
                    int k = code.lastIndexOf('.');
                    if (k < 0) {
                        k = code.length();
                        for (; k > 0; k--) {
                            char c = code.charAt(k - 1);
                            if (!(c == '_' || c == ' ' || c == '.')) break;
                        }
                    }
                    int _len = len;
                    if (len >= 2 && code.charAt(len - 1) == ')' && code.charAt(len - 2) == '_') _len--;
                    k = k >= 0 && _len > k ? _len - k : 0;
                    ks[i] = (k << 1) | (comma ? 1 : 0);
                }
            }

            if (!isDate) {
                int max = Arrays.stream(ks).max().orElse(0);
                wc = max << 1;
                wc |= ((int) (s * 10000 + 0.5)) << 8;
            }
            return wc;
        });
    }

    @Override
    public int hashCode() {
        return code != null ? code.hashCode() : 0;
    }

    @Override
    public boolean equals(Object o) {
        if (o instanceof NumFmt) {
            NumFmt other = (NumFmt) o;
            return Objects.equals(other.code, code);
        }
        return false;
    }

    @Override
    public String toString() {
        return "id: " + id + ", code: " + code;
    }

    public Element toDom(Element root) {
        if (StringUtil.isEmpty(code)) return root; // Build in style
        return root.addElement(StringUtil.lowFirstKey(getClass().getSimpleName()))
            .addAttribute("formatCode", code)
            .addAttribute("numFmtId", String.valueOf(id));
    }

    public static List<NumFmt> domToNumFmt(Element root) {
        // Number format
        Element ele = root.element("numFmts");
        // Break if there don't contains 'numFmts' tag
        if (ele == null) {
            return new ArrayList<>();
        }
        List<Element> sub = ele.elements();
        List<NumFmt> numFmts = new ArrayList<>(sub.size());
        for (Element e : sub) {
            String id = getAttr(e, "numFmtId"), code = getAttr(e, "formatCode");
            numFmts.add(new NumFmt(Integer.parseInt(id), code));
        }
        // Sort by id
        numFmts.sort(Comparator.comparingInt(NumFmt::getId));
        return numFmts;
    }

    @Override
    public int compareTo(NumFmt o) {
        return id - o.id;
    }
}

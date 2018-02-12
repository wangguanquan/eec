package net.cua.export.entity.e7.style;

import net.cua.export.util.StringUtil;
import org.dom4j.Element;

/**
 * To create a custom number format, you start by selecting one of the built-in number formats as a starting point.
 * You can then change any one of the code sections of that format to create your own custom number format.

 A number format can have up to four sections of code, separated by semicolons.
 These code sections define the format for positive numbers, negative numbers, zero values, and text, in that order.

 <POSITIVE>;<NEGATIVE>;<ZERO>;<TEXT>

 For example, you can use these code sections to create the following custom format:

 [Blue]#,##0.00_);[Red](#,##0.00);0.00;"sales "@

 You do not have to include all code sections in your custom number format.
 If you specify only two code sections for your custom number format,
 the first section is used for positive numbers and zeros, and the second section is used for negative numbers.
 If you specify only one code section, it is used for all numbers.
 If you want to skip a code section and include a code section that follows it,
 you must include the ending semicolon for the section that you skip.
 *
 * https://support.office.com/en-us/article/create-or-delete-a-custom-number-format-78f2a361-936b-4c03-8772-09fab54be7f4
 * Created by guanquan.wang at 2018-02-06 08:51
 */
public class NumFmt {
    private String code;
    private int id = -1;
    private NumFmt(){}

    public NumFmt(String code) {
        this.code = code;
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
     * 内置format
     * @param id
     * @return
     */
    public static final NumFmt valueOf(int id) {
        return new NumFmt().setId(id);
    }

    @Override
    public int hashCode() {
        return code != null ? code.hashCode() : 0;
    }

    @Override
    public boolean equals(Object o) {
        if (o instanceof NumFmt) {
            NumFmt other = (NumFmt) o;
            return other.code != null ? other.code.equals(code) : other.code == code;
        }
        return false;
    }

    public Element toDom4j(Element root) {
        if (StringUtil.isEmpty(code)) return root; // 内置format不用重复输出
        return root.addElement(StringUtil.lowFirstKey(getClass().getSimpleName()))
                .addAttribute("formatCode", code)
                .addAttribute("numFmtId", String.valueOf(id));
    }
}

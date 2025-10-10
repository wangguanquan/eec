/*
 * Copyright (c) 2017-2023, guanquan.wang@hotmail.com All Rights Reserved.
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


package org.ttzero.excel.validation;

import org.dom4j.Element;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.ttzero.excel.manager.Const;
import org.ttzero.excel.reader.CrossDimension;
import org.ttzero.excel.reader.Dimension;
import org.ttzero.excel.util.StringUtil;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.stream.Collectors;

import static org.ttzero.excel.reader.Row.testNumberType;

/**
 * 数据验证
 *
 * @author guanquan.wang at 2022-08-17 20:05:42
 */
public abstract class Validation {
    /**
     * LOGGER
     */
    static final Logger LOGGER = LoggerFactory.getLogger(Validation.class);
    /**
     * 允许为空
     */
    public boolean allowBlank = true;
    /**
     * 显示下拉框
     */
    public boolean showInputMessage = true;
    /**
     * 显示提示信息
     */
    public boolean showErrorMessage = true;
    /**
     * 作用范围（支持多个范围值）
     */
    public List<Dimension> sqrefList = new ArrayList<>();
    /**
     * 操作符，不指定时默认between
     */
    public Operator operator;
    /**
     * 引用其它工作表维度
     */
    public CrossDimension referer;
    /**
     * 提示, 提示的标题
     */
    public String prompt, promptTitle;
    /**
     * 错误消息, 错误警报的标题, 数据验证的错误警报样式
     */
    public String error, errorTitle;
    /**
     * 错误警报样式(stop:提供"取消/重试"选项，必须修改正确才能离开单元格，warning:提供"是/否"选项，information:只提供"确定"按钮)
     */
    public ErrorStyle errorStyle;
    /**
     * 数据校验类型
     *
     * @return 数据验类型，包含 {@code list}序列, {@code whole}整数, {@code date}日期, {@code time}时间, {@code textLength}文本长度
     */
    public abstract String getType();

    /**
     * 校验内容
     *
     * @return 验证对象转xml文本
     */
    public abstract String validationFormula();

    /**
     * 是否为扩展节点
     *
     * @return {@code true} 扩展节点
     */
    public boolean isExtension() {
        return referer != null && referer.isCrossSheet();
    }

    /**
     * 设置作用域
     *
     * @param sqref 作用域
     * @return 当前数据验证
     */
    public Validation dimension(Dimension sqref) {
        if (!sqrefList.contains(sqref)) sqrefList.add(sqref);
        return this;
    }

    /**
     * 提示提示
     *
     * @param prompt 提示词
     * @return 当前数据验证
     */
    public Validation prompt(String prompt) {
        return prompt(null, prompt);
    }

    /**
     * 提示提示和标题
     *
     * @param promptTitle 提示词标题
     * @param prompt 提示词
     * @return 当前数据验证
     */
    public Validation prompt(String promptTitle, String prompt) {
        this.promptTitle = promptTitle;
        this.prompt = prompt;
        return this;
    }

    /**
     * 设置跨工作表维度引用
     *
     * @param referer 跨工作表维度引用
     * @return 当前数据验证
     */
    public Validation referer(CrossDimension referer) {
        this.referer = referer;
        return this;
    }

    /**
     * 设置数据校验NG时提示错误消息
     *
     * @param error 错误消息
     * @return 当前数据验证
     */
    public Validation error(String error) {
        return error(null, error);
    }

    /**
     * 设置数据校验NG时提示错误消息
     *
     * @param error 错误消息
     * @param errorStyle 错误警报样式
     * @return 当前数据验证
     */
    public Validation error(String error, ErrorStyle errorStyle) {
        return error(null, error, errorStyle);
    }

    /**
     * 设置数据校验NG时提示错误消息和标题
     *
     * @param errorTitle 错误警报的标题
     * @param error 错误消息
     * @return 当前数据验证
     */
    public Validation error(String errorTitle, String error) {
        this.errorTitle = errorTitle;
        this.error = error;
        return this;
    }

    /**
     * 设置数据校验NG时提示错误消息和标题
     *
     * @param errorTitle 错误警报的标题
     * @param error 错误消息
     * @param errorStyle 错误警报样式
     * @return 当前数据验证
     */
    public Validation error(String errorTitle, String error, ErrorStyle errorStyle) {
        this.errorTitle = errorTitle;
        this.error = error;
        this.errorStyle = errorStyle;
        return this;
    }

    protected String getSqrefStr() {
        if (sqrefList != null && !sqrefList.isEmpty()) {
            StringBuilder buf = new StringBuilder(sqrefList.get(0).toString());
            for (int i = 1; i < sqrefList.size(); i++) {
                buf.append(' ').append(sqrefList.get(i));
            }
            return buf.toString();
        }
        return StringUtil.EMPTY;
    }

    @Override
    public String toString() {
        return "<" + (isExtension() ? "x14:" : "" ) + "dataValidation type=\"" + getType()
            + (errorStyle == null || errorStyle == ErrorStyle.stop ? "" : "\" errorStyle=\"" + errorStyle)
            + (operator == null || operator == Operator.between ? "" : "\" operator=\"" + operator)
            + "\" allowBlank=\"" + (allowBlank ? 1 : 0)
            + "\" showInputMessage=\"" + (showInputMessage ? 1 : 0)
            + "\" showErrorMessage=\"" + (showErrorMessage ? 1 : 0)
            + (StringUtil.isEmpty(errorTitle) ? "" : "\" errorTitle=\"" + StringUtil.escapeString(errorTitle))
            + (StringUtil.isEmpty(error) ? "" : "\" error=\"" + StringUtil.escapeString(error))
            + (StringUtil.isEmpty(promptTitle) ? "" : "\" promptTitle=\"" + StringUtil.escapeString(promptTitle))
            + (StringUtil.isEmpty(prompt) ? "" : "\" prompt=\"" + StringUtil.escapeString(prompt))
            + (isExtension() ? "\">" : "\" sqref=\"" + getSqrefStr() + "\">")
            + validationFormula()
            + "</" + (isExtension() ? "x14:" : "" ) + "dataValidation>";
    }

    protected void parseAttribute(Element e, boolean isExt) {
        this.allowBlank = "1".equals(e.attributeValue("allowBlank"));
        this.showInputMessage = "1".equals(e.attributeValue("showInputMessage"));
        this.showErrorMessage = "1".equals(e.attributeValue("showErrorMessage"));
        this.prompt = e.attributeValue("prompt");
        this.errorTitle = e.attributeValue("errorTitle");
        this.error = e.attributeValue("error");
        String tmp = isExt ? e.elementText("sqref") : e.attributeValue("sqref");
        if (StringUtil.isNotEmpty(tmp)) {
            this.sqrefList.addAll(Arrays.stream(tmp.split(" ")).filter(StringUtil::isNotEmpty).map(Dimension::of).collect(Collectors.toList()));
        }
        tmp = e.attributeValue("operator");
        if (StringUtil.isNotEmpty(tmp)) this.operator = Validation.Operator.valueOf(tmp);
        tmp = e.attributeValue("errorStyle");
        if (StringUtil.isNotEmpty(tmp)) {
            try {
                this.errorStyle = ErrorStyle.valueOf(tmp);
            } catch (Exception ex) {
                // Ignore
            }
        }
    }

    public static List<Validation> domToValidation(Element e) {
        List<Element> sub =  e.elements("dataValidation");
        if (sub == null || sub.isEmpty()) return null;
        final boolean isExt = Const.NAMESPACE.X14.equals(e.getNamespace().getURI());
        List<Validation> validations = new ArrayList<>(sub.size());
        for (Element o : sub) {
            String type = o.attributeValue("type");
            Validation val = null;
            if (type != null) {
                switch (type) {
                    case "list"      : val = new ListValidation<String>(); break;
                    case "time"      : val = new TimeValidation();         break;
                    case "date"      : val = new DateValidation();         break;
                    case "textLength": val = new TextLengthValidation();   break;
                    case "whole"     : val = new WholeValidation();        break;
                    case "decimal"   : val = new DecimalValidation();      break;
                    default:
                }
            }
            if (val != null) {
                val.parseAttribute(o, isExt);
                validations.add(val);
            }
        }
        return validations;
    }

    /**
     * 检查节点类型
     *
     * @param txt 文本
     * @return 0:无 1:字符串 2:坐标 3:跨工作表坐标 4:公式 5:数字
     */
    static int testValueType(String txt) {
        if (StringUtil.isEmpty(txt)) return 0;
        int t, len = txt.length();
        if (len >= 2 && txt.charAt(0) == '"' && txt.charAt(len - 1) == '"') t = 1;
        else if (testNumberType(txt.toCharArray(), 0, txt.length()) > 0) t = 5;
        else if (txt.indexOf('(') > 0 && txt.charAt(len - 1) == ')') t = 4;
        else {
            int i = txt.indexOf('!');
            String r = i > 0 && i < len - 2 ? txt.substring(i + 1) : txt;
            try {
                Dimension.of(r.replace("$", ""));
                t = i > 0 ? 3 : 2;
            } catch (Exception e) {
                t = 1;
            }
        }
        return t;
    }

    public enum Operator {
        between, notBetween, equal, notEqual, greaterThan, lessThan, greaterThanOrEqual, lessThanOrEqual
    }

    public enum ErrorStyle {
        stop, warning, information
    }
}

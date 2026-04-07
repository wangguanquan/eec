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
import org.ttzero.excel.reader.CrossDimension;
import org.ttzero.excel.reader.Dimension;
import org.ttzero.excel.util.StringUtil;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

/**
 * 序列验证，限定单元格的值只能在序列中选择，当可选序列值文本合计长度超过255时将转为引用序列
 *
 * @author guanquan.wang at 2022-08-17 20:05:42
 */
public class ListValidation<T> extends Validation {
    /**
     * 可选的序列值
     */
    public List<T> options;
    /**
     * 级联序列
     */
    public List<CascadeList<T>> cascadeList;
    /**
     * 级联INDIRECT函数
     */
    public String indirect;

    public ListValidation<T> in(List<T> options) {
        this.options = options;
        return this;
    }

    @SafeVarargs
    public final ListValidation<T> in(T... options) {
        this.options = Arrays.asList(options);
        return this;
    }

    public ListValidation<T> in(Dimension referer) {
        return in(null, referer);
    }

    public ListValidation<T> in(String otherSheetName, Dimension referer) {
        this.referer = new CrossDimension(otherSheetName, referer);
        return this;
    }

    public static <T> ListValidation<T> in(Dimension sqref, List<T> options) {
        ListValidation<T> lv = new ListValidation<>();
        lv.options = options;
        lv.dimension(sqref);
        return lv;
    }

    public ListValidation<T> addCascadeList(Dimension sqref, Map<T, List<T>> subList) {
        if (this.cascadeList == null) {
            this.cascadeList = new ArrayList<>();
        }
        this.cascadeList.add(new CascadeList<>(sqref, subList));
        return this;
    }

    public CascadeList<T> getCascadeList(int level) {
        return level >= 1 && level <= getCascadeSize() ? cascadeList.get(level - 1) : null;
    }

    public int getCascadeSize() {
        return cascadeList != null ? cascadeList.size() : 0;
    }

    @Override
    public String getType() {
        return "list";
    }

    @Override
    public String validationFormula() {
        String val;
        if (isExtension()) {
            val = "<x14:formula1><xm:f>" + referer + "</xm:f></x14:formula1><xm:sqref>" + getSqrefStr() + "</xm:sqref>";
        } else if (options != null) {
            val = "<formula1>\"" + options.stream().map(String::valueOf).map(StringUtil::escapeString).collect(Collectors.joining(",")) + "\"</formula1>";
        } else if (referer != null) {
            val = "<formula1>" + referer + "</formula1>";
        } else if (StringUtil.isNotEmpty(indirect)) {
            val = "<formula1>INDIRECT(" + indirect + ")</formula1>";
        } else {
            val = "<formula1>\"\"</formula1>";
        }
        return val;
    }

    public static class CascadeList<T> {
        public Map<T, List<T>> options;
        public Dimension sqref;

        CascadeList(Dimension sqref, Map<T, List<T>> options) {
            this.options = options;
            this.sqref = sqref;
        }
    }

    @Override
    protected void parseAttribute(Element e, boolean isExt) {
        super.parseAttribute(e, isExt);
        Element formula1 = e.element("formula1");
        // 非法
        if (formula1 == null) return;
        String txt = isExt ? formula1.elementText("f") : formula1.getText();
        // 0:无 1:字符串 2:坐标 3:跨工作表坐标 4:公式 5:数字
        int type = testValueType(txt);
        switch (type) {
            case 1:
                this.options = (List<T>) Arrays.asList(txt.substring(1, txt.length() - 1).split(","));
                break;
            case 5:
                this.options = (List<T>) Collections.singletonList(txt);
                break;
            case 2:
            case 3:
                this.referer = CrossDimension.of(txt);
                break;
            case 4:
                // FIXME 公式只支持INDIRECT,其它暂时不支持，示例:SUM(A1:A100)
                if (txt.toUpperCase().startsWith("INDIRECT(")) {
                    this.indirect = txt.substring(9, txt.length() - 1);
                }
                else {
                    LOGGER.warn("Unsupported formula value[{}]", txt);
                }
                break;
            default:
                LOGGER.warn("Unsupported formula value[{}]", txt);
        }
    }
}

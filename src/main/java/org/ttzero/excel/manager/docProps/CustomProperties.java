/*
 * Copyright (c) 2017-2024, guanquan.wang@hotmail.com All Rights Reserved.
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


package org.ttzero.excel.manager.docProps;

import org.dom4j.Element;
import org.dom4j.Namespace;
import org.dom4j.QName;
import org.ttzero.excel.manager.TopNS;
import org.ttzero.excel.util.DateUtil;
import org.ttzero.excel.util.StringUtil;

import java.util.Date;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.Map;

/**
 * 自定义属性
 *
 * <p>注意：只支持{@code "文本"}、{@code "数字"}、{@code "日期"}以及{@code "布尔值"}，其它数据类型将使用{@code toString}强转换为文本</p>
 *
 * @author guanquan.wang
 * @since 2024-09-19
 */
@TopNS(prefix = {"", "vt"}, uri = {"http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
    , "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"}, value = "Properties")
public class CustomProperties extends XmlEntity {
    /**
     * 自定义属性的GUID值{D5CDD505-2E9C-101B-9397-08002B2CF9AE}
     */
    public static final String FORMAT_ID = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}";
    /**
     * 自定义属性
     * key: 属性名
     * value: v1: 属性值 v2: 值类型
     */
    private final Map<String, Tuple2<Object, Integer>> properties;

    public CustomProperties() {
        this.properties = new LinkedHashMap<>();
    }

    /**
     * 将指定的键值对添加到属性集合中
     *
     * @param key 属性的键名
     * @param value 属性的值
     */
    public void put(String key, Object value) {
        check(key, value);
        properties.put(key, toProValue(value));
    }

    /**
     * 将指定的属性集合添加到当前对象中
     *
     * @param properties 属性名称和值的映射表
     */
    public void putAll(Map<String, Object> properties) {
        for (Map.Entry<String, Object> entry : properties.entrySet()) {
            check(entry.getKey(), entry.getValue());
            this.properties.put(entry.getKey(), toProValue(entry.getValue()));
        }
    }

    /**
     * 移除指定属性
     *
     * @param key 指定需要移除的Key
     * @return 如果Key存在则返回对应的值否则返回 {@code null}
     */
    public Object remove(String key) {
        Tuple2<Object, Integer> v = properties.remove(key);
        return v != null ? v.v1 : null;
    }

    /**
     * 获取所有自定义属性的副本
     *
     * @return 自定义属性列表
     */
    public Map<String, Object> getAllProperties() {
        Map<String, Object> result = new HashMap<>(properties.size());
        for (Map.Entry<String, Tuple2<Object, Integer>> entry : properties.entrySet()) {
            result.put(entry.getKey(), entry.getValue().v1);
        }
        return result;
    }

    /**
     * 测试是否包含自定义属性
     *
     * @return true: 包含自定义属性
     */
    public boolean hasProperty() {
        return !properties.isEmpty();
    }

    @Override
    void toDom(Element root, Map<String, Namespace> namespaceMap) {
        int id = 2; // beginning pid
        Namespace vt = namespaceMap.get("vt");
        for (Map.Entry<String, Tuple2<Object, Integer>> entry : properties.entrySet()) {
            Element property = root.addElement("property").addAttribute("fmtid", FORMAT_ID)
                .addAttribute("pid", Integer.toString(id++)).addAttribute("name", entry.getKey());
            Tuple2<Object, Integer> val = entry.getValue();
            switch (val.v2) {
                case 0: property.addElement(QName.get("lpwstr", vt)).addText(val.v1.toString());                   break;
                case 1: property.addElement(QName.get("filetime", vt)).addText(DateUtil.toTString((Date) val.v1)); break;
                case 2: property.addElement(QName.get("i4", vt)).addText(val.v1.toString());                       break;
                case 3: property.addElement(QName.get("r8", vt)).addText(val.v1.toString());                       break;
                case 4: property.addElement(QName.get("bool", vt)).addText(val.v1.toString());                     break;
                default:
            }
        }
    }

    /**
     * 检查属性的合法性
     *
     * @param key 属性名
     * @param val 属性值
     */
    protected static void check(String key, Object val) {
        if (StringUtil.isEmpty(key))
            throw new IllegalArgumentException("Property name is required.");
        if (key.length() > 256)
            throw new IllegalArgumentException("Property name is too long, max=256 current=" + key.length());
        if (val == null)
            throw new IllegalArgumentException("Property value is required.");
    }

    /**
     * 将属性值转换为可识别的结果
     *
     * @param val 外部属性值
     * @return 内部属性值 v1: 原值 v2：类型
     */
    protected static Tuple2<Object, Integer> toProValue(Object val) {
        int t;
        if (val instanceof String) t = 0;
        else if (val instanceof Date) t = 1;
        else if (val instanceof Integer || val instanceof Short) t = 2;
        else if (val instanceof Long || val instanceof Double || val instanceof Float) t = 3;
        else if (val instanceof Boolean) t = 4;
        else t = 0;
        return Tuple2.of(val, t);
    }
}

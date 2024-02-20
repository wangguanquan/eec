/*
 * Copyright (c) 2017, guanquan.wang@yandex.com All Rights Reserved.
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
import org.ttzero.excel.manager.NS;
import org.ttzero.excel.manager.TopNS;
import org.ttzero.excel.util.DateUtil;

import java.lang.reflect.Field;
import java.util.Date;
import java.util.Map;

/**
 * 文档属性，指定主题，作者和关键词等信息，可以通过鼠标右建-&gt;详细属性查看这些内容
 *
 * @author guanquan.wang on 2017/9/18.
 */
@TopNS(prefix = {"dc", "cp", "dcterms"}
    , uri = {"http://purl.org/dc/elements/1.1/"
    , "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
    , "http://purl.org/dc/terms/"}, value = "cp:coreProperties")
public class Core extends XmlEntity {
    /**
     * 标题
     */
    @NS("dc")
    private String title;
    /**
     * 主题
     */
    @NS("dc")
    private String subject;
    /**
     * 作者，默认使用当前系统登陆用户名
     */
    @NS("dc")
    private String creator;
    /**
     * 描述
     */
    @NS("dc")
    private String description;
    /**
     * 关键词，多个关键词使用分号{@code ';'}分隔
     */
    @NS("cp")
    private String keywords;
    /**
     * 最后修改人
     */
    @NS("cp")
    private String lastModifiedBy;
    /**
     * 版本，默认使用EEC版本号
     */
    @NS("cp")
    private String version;
    /**
     * 修订版本号
     */
    @NS("cp")
    private String revision;
    /**
     * 分类，多个分类词使用分号{@code ';'}分隔
     */
    @NS("cp")
    private String category;
    /**
     * 创建时间
     */
    @NS("dcterms")
    private Date created;
    /**
     * 修改时间
     */
    @NS("dcterms")
    private Date modified;

    public void setTitle(String title) {
        this.title = title;
    }

    public void setSubject(String subject) {
        this.subject = subject;
    }

    public void setCreator(String creator) {
        this.creator = creator;
    }

    public void setDescription(String description) {
        this.description = description;
    }

    public void setKeywords(String keywords) {
        this.keywords = keywords;
    }

    public void setLastModifiedBy(String lastModifiedBy) {
        this.lastModifiedBy = lastModifiedBy;
    }

    public void setVersion(String version) {
        this.version = version;
    }

    public void setRevision(String revision) {
        this.revision = revision;
    }

    public void setCategory(String category) {
        this.category = category;
    }

    public void setCreated(Date created) {
        this.created = created;
    }

    public void setModified(Date modified) {
        this.modified = modified;
    }

    public String getTitle() {
        return title;
    }

    public String getSubject() {
        return subject;
    }

    public String getCreator() {
        return creator;
    }

    public String getDescription() {
        return description;
    }

    public String getKeywords() {
        return keywords;
    }

    public String getLastModifiedBy() {
        return lastModifiedBy;
    }

    public String getVersion() {
        return version;
    }

    public String getRevision() {
        return revision;
    }

    public String getCategory() {
        return category;
    }

    public Date getCreated() {
        return created;
    }

    public Date getModified() {
        return modified;
    }

    @Override
    void toDom(Element rootElement, Map<String, Namespace> namespaceMap) {
        Field[] fields = getClass().getDeclaredFields();
        for (Field field : fields) {
            NS ns = field.getDeclaredAnnotation(NS.class);
            if (ns == null) continue;
            Object o;
            try {
                o = field.get(this);
            } catch (IllegalAccessException e) {
                continue;
            }
            if (o == null) continue;
            if ("dcterms".equals(ns.value()) && (o instanceof java.util.Date)) {
                Element e = rootElement.addElement(QName.get(field.getName(), namespaceMap.get(ns.value()))).addText(DateUtil.toTString((Date) o));
                e.addAttribute(QName.get("type", Namespace.get("xsi", "http://www.w3.org/2001/XMLSchema-instance")), "dcterms:W3CDTF");
            } else rootElement.addElement(QName.get(field.getName(), namespaceMap.get(ns.value()))).addText(o.toString());
        }
    }
}

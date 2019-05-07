/*
 * Copyright (c) 2019, guanquan.wang@yandex.com All Rights Reserved.
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

package cn.ttzero.excel.manager.docProps;


import cn.ttzero.excel.annotation.Attr;
import cn.ttzero.excel.annotation.NS;
import cn.ttzero.excel.annotation.TopNS;

import java.util.Date;

/**
 * Created by guanquan.wang on 2017/9/18.
 */
@TopNS(prefix = {"dc", "cp", "dcterms"}
    , uri = {"http://purl.org/dc/elements/1.1/"
    , "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
    , "http://purl.org/dc/terms/"}, value = "cp:coreProperties")
public class Core extends XmlEntity {
    @NS("dc")
    private String title;   // 标题
    @NS("dc")
    private String subject; // 主题
    @NS("dc")
    private String creator; // 作者
    @NS("dc")
    private String description; // 描述
    @NS("cp")
    private String keywords;        // 标记
    @NS("cp")
    private String lastModifiedBy;  // 最后一次保存者
    @NS("cp")
    private String version;         // 版本号
    @NS("cp")
    private String revision;        // 修订号
    @NS("cp")
    private String category;        // 类别
    @NS("dcterms")
    @Attr(name = "type", value = "dcterms:W3CDTF", namespace = @NS(value = "xsi", uri = "http://www.w3.org/2001/XMLSchema-instance"))
    private Date created;         // 创建时间
    @NS("dcterms")
    @Attr(name = "type", value = "dcterms:W3CDTF", namespace = @NS(value = "xsi", uri = "http://www.w3.org/2001/XMLSchema-instance"))
    private Date modified;       // 最后更新时间


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
}

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


import org.ttzero.excel.annotation.Attr;
import org.ttzero.excel.annotation.NS;
import org.ttzero.excel.annotation.TopNS;

import java.util.Date;

/**
 * @author guanquan.wang on 2017/9/18.
 */
@TopNS(prefix = {"dc", "cp", "dcterms"}
    , uri = {"http://purl.org/dc/elements/1.1/"
    , "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
    , "http://purl.org/dc/terms/"}, value = "cp:coreProperties")
public class Core extends XmlEntity {
    /**
     * Title of workbook(null able)
     */
    @NS("dc")
    private String title;
    /**
     * Subject of workbook(null able)
     */
    @NS("dc")
    private String subject;
    /**
     * Author, default get the system login user
     */
    @NS("dc")
    private String creator;
    /**
     * Description(null able)
     */
    @NS("dc")
    private String description;
    /**
     * List keyword about this workbook, Multiple keywords are separated by ';'(null able)
     */
    @NS("cp")
    private String keywords;
    /**
     * The last modify user(null able)
     */
    @NS("cp")
    private String lastModifiedBy;
    /**
     * The file version(null able)
     */
    @NS("cp")
    private String version;
    /**
     * The file reversion(null able)
     */
    @NS("cp")
    private String revision;
    /**
     * Specify category about this workbook, Multiple keywords are separated by ';'(null able)
     */
    @NS("cp")
    private String category;
    /**
     * Create time(notnull)
     */
    @NS("dcterms")
    @Attr(name = "type", value = "dcterms:W3CDTF", namespace = @NS(value = "xsi", uri = "http://www.w3.org/2001/XMLSchema-instance"))
    private Date created;
    /**
     * The last modify time(notnull)
     */
    @NS("dcterms")
    @Attr(name = "type", value = "dcterms:W3CDTF", namespace = @NS(value = "xsi", uri = "http://www.w3.org/2001/XMLSchema-instance"))
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
}

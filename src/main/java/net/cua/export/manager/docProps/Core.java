package net.cua.export.manager.docProps;


import net.cua.export.annotation.Attr;
import net.cua.export.annotation.NS;
import net.cua.export.annotation.TopNS;

import java.util.Date;

/**
 * Created by wanggq on 2017/9/18.
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
    private Date created = new Date();         // 创建时间
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

}

/*
 * Copyright (c) 2019-2021, guanquan.wang@yandex.com All Rights Reserved.
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

package org.ttzero.excel.reader;

import org.ttzero.excel.manager.ExcelType;
import org.ttzero.excel.manager.docProps.App;
import org.ttzero.excel.manager.docProps.Core;
import org.ttzero.excel.util.DateUtil;
import org.ttzero.excel.util.StringUtil;

import java.util.Date;

/**
 * The file basic information
 *
 * Create by guanquan.wang at 2019-04-15 16:00
 */
public class AppInfo {
    private final App app;
    private final Core core;
    private final ExcelType type;

    AppInfo(App app, Core core, ExcelType type) {
        this.app = app;
        this.core = core;
        this.type = type;
    }

    /**
     * The app name for file creation
     *
     * @return the app name
     */
    public String getApplication() {
        return app.getApplication();
    }

    /**
     * About company info
     *
     * @return company value or null
     */
    public String getCompany() {
        return app.getCompany();
    }

    /**
     * The app version for file creation
     *
     * @return the app version not null
     */
    public String getAppVersion() {
        return app.getAppVersion();
    }

    /**
     * The content's title
     *
     * @return title value or null
     */
    public String getTitle() {
        return core.getTitle();
    }

    /**
     * The content's subject
     *
     * @return subject value or null
     */
    public String getSubject() {
        return core.getSubject();
    }

    /**
     * The file creator name
     *
     * @return name of creator or null
     */
    public String getCreator() {
        return core.getCreator();
    }

    /**
     * The content's short description
     *
     * @return the desc value or null
     */
    public String getDescription() {
        return core.getDescription();
    }

    /**
     * The content keyword, join with ','
     *
     * @return keyword value or null
     */
    public String getKeywords() {
        return core.getKeywords();
    }

    /**
     * Returns the last modified user
     *
     * @return the modified user or null
     */
    public String getLastModifiedBy() {
        return core.getLastModifiedBy();
    }

    /**
     * Returns the App version
     *
     * @return the excel app version or the eec version if it used eec
     */
    public String getVersion() {
        return core.getVersion();
    }

    /**
     * Returns the file revision
     *
     * @return revision or null
     */
    public String getRevision() {
        return core.getRevision();
    }

    /**
     * Returns the content's category, the value join with ','
     *
     * @return category value or null
     */
    public String getCategory() {
        return core.getCategory();
    }

    /**
     * Returns the create time
     *
     * @return the file create date or null
     */
    public Date getCreated() {
        return core.getCreated();
    }

    /**
     * Returns the last modified
     *
     * @return last modified date or null
     */
    public Date getModified() {
        return core.getModified();
    }

    /**
     * Returns the file type
     *
     * @return the file {@link ExcelType}
     */
    public ExcelType getType() {
        return type;
    }

    @Override
    public String toString() {
        return "Type: " + type
            + "Application: " + getApplication()
            + " Company: " + getCompany()
            + " AppVersion: " + getAppVersion()
            + " Title: " + getTitle()
            + " Subject: " + getSubject()
            + " Creator: " + getCreator()
            + " Description: " + getDescription()
            + " Keywords: " + getKeywords()
            + " LastModifiedBy: " + getLastModifiedBy()
            + " Version: " + getVersion()
            + " Revision: " + getRevision()
            + " Category: " + getCategory()
            + " Created: " + (getCreated() != null ? DateUtil.toString(getCreated()) : StringUtil.EMPTY)
            + " Modified: " + (getModified() != null ? DateUtil.toString(getModified()) : StringUtil.EMPTY)
            ;
    }
}

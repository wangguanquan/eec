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

package cn.ttzero.excel.reader;

import cn.ttzero.excel.manager.docProps.App;
import cn.ttzero.excel.manager.docProps.Core;
import cn.ttzero.excel.util.DateUtil;
import cn.ttzero.excel.util.StringUtil;

import java.util.Date;

/**
 * Create by guanquan.wang at 2019-04-15 16:00
 */
public class AppInfo {
    private App app;
    private Core core;

    AppInfo(App app, Core core) {
        this.app = app;
        this.core = core;
    }

    public String getApplication() {
        return app.getApplication();
    }

    public String getManager() {
        return app.getManager();
    }

    public String getCompany() {
        return app.getCompany();
    }

    public String getAppVersion() {
        return app.getAppVersion();
    }

    public String getTitle() {
        return core.getTitle();
    }

    public String getSubject() {
        return core.getSubject();
    }

    public String getCreator() {
        return core.getCreator();
    }

    public String getDescription() {
        return core.getDescription();
    }

    public String getKeywords() {
        return core.getKeywords();
    }

    public String getLastModifiedBy() {
        return core.getLastModifiedBy();
    }

    public String getVersion() {
        return core.getVersion();
    }

    public String getRevision() {
        return core.getRevision();
    }

    public String getCategory() {
        return core.getCategory();
    }

    public Date getCreated() {
        return core.getCreated();
    }

    public Date getModified() {
        return core.getModified();
    }

    @Override
    public String toString() {
        return "Application: " + getApplication()
            + " Manager: " + getManager()
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

/*
 * Copyright (c) 2017-2019, guanquan.wang@yandex.com All Rights Reserved.
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

import org.ttzero.excel.manager.docProps.App;
import org.ttzero.excel.manager.docProps.Core;
import org.ttzero.excel.util.DateUtil;
import org.ttzero.excel.util.ReflectUtil;

import java.lang.reflect.Method;
import java.util.Date;
import java.util.StringJoiner;

/**
 * The file basic information
 *
 * @author guanquan.wang at 2019-04-15 16:00
 */
public class AppInfo {
    private final App app;
    private final Core core;

    AppInfo(App app, Core core) {
        this.app = app;
        this.core = core;
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

    @Override
    public String toString() {
        StringJoiner joiner = new StringJoiner(System.lineSeparator());
        try {
            Method[] methods = ReflectUtil.listReadMethods(getClass(),
                method -> method.getName().startsWith("get")
                    || method.getReturnType() == boolean.class && method.getName().startsWith("is"));
            for (Method method : methods) {
                Object o = method.invoke(this);
                if (o != null) {
                    joiner.add(method.getReturnType() == boolean.class ? method.getName().substring(2)
                        : method.getName().substring(3) + ": "
                        + (o instanceof Date ? DateUtil.toString((Date)o) : o.toString()));
                }
            }
        } catch (Exception e) {
            // Ignore
        }
        return joiner.toString();
    }
}

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
import org.ttzero.excel.util.ReflectUtil;

import java.lang.reflect.Method;
import java.util.Date;
import java.util.StringJoiner;

import static org.ttzero.excel.util.DateUtil.toDateTimeString;

/**
 * Excel文件基础信息包含作者、日期等信息，在windows操作系统上使用鼠标右键-&gt;属性-&gt;详细信息查看
 *
 * @author guanquan.wang at 2019-04-15 16:00
 */
public class AppInfo {
    /**
     * App属性
     */
    private final App app;
    /**
     * 文档详细属性
     */
    private final Core core;

    AppInfo(App app, Core core) {
        this.app = app;
        this.core = core;
    }

    /**
     * 获取Excel由哪个App生成或打开
     *
     * @return 属性-详细属性-程序名称
     */
    public String getApplication() {
        return app.getApplication();
    }

    /**
     * 获取公司名
     *
     * @return 属性-详细属性-公司
     */
    public String getCompany() {
        return app.getCompany();
    }

    /**
     * 获取App版号
     *
     * @return 属性-详细属性-版本号
     */
    public String getAppVersion() {
        return app.getAppVersion();
    }

    /**
     * 获取标题
     *
     * @return 属性-详细属性-标题
     */
    public String getTitle() {
        return core.getTitle();
    }

    /**
     * 获取主题
     *
     * @return 属性-详细属性-主题
     */
    public String getSubject() {
        return core.getSubject();
    }

    /**
     * 获取作者
     *
     * @return 属性-详细属性-作者
     */
    public String getCreator() {
        return core.getCreator();
    }

    /**
     * 获取备注
     *
     * @return 属性-详细属性-备注
     */
    public String getDescription() {
        return core.getDescription();
    }

    /**
     * 获取标记，多个标记使用{@code ','}逗号分隔
     *
     * @return 属性-详细属性-标记
     */
    public String getKeywords() {
        return core.getKeywords();
    }

    /**
     * 获取最后一次保存者
     *
     * @return 属性-详细属性-最后一次保存者
     */
    public String getLastModifiedBy() {
        return core.getLastModifiedBy();
    }

    /**
     * 获取版本号
     *
     * @return 属性-详细属性-版本号
     */
    public String getVersion() {
        return core.getVersion();
    }

    /**
     * 获取修订号
     *
     * @return 属性-详细属性-修订号
     */
    public String getRevision() {
        return core.getRevision();
    }

    /**
     * 获取类别，多个类别使用{@code ','}逗号分隔
     *
     * @return 属性-详细属性-类别
     */
    public String getCategory() {
        return core.getCategory();
    }

    /**
     * 获取创建时间
     *
     * @return 属性-详细属性-创建时间
     */
    public Date getCreated() {
        return core.getCreated();
    }

    /**
     * 获取修改时间
     *
     * @return 属性-详细属性-修改时间
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
                        + (o instanceof Date ? toDateTimeString((Date) o) : o.toString()));
                }
            }
        } catch (Exception e) {
            // Ignore
        }
        return joiner.toString();
    }
}

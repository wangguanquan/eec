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


import org.ttzero.excel.manager.Attr;
import org.ttzero.excel.manager.NS;
import org.ttzero.excel.manager.TopNS;

import java.util.ArrayList;
import java.util.List;

import static org.ttzero.excel.util.StringUtil.isEmpty;

/**
 * App属性，除{@code company}属性外其余属性均由{@link org.ttzero.excel.entity.Workbook}生成，
 * 外部不要随意修改否则将导致不可预期的异常。
 *
 * @author guanquan.wang on 2017/9/21.
 */
@TopNS(prefix = {"vt", ""}, uri = {"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
    , "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"}, value = "Properties")
public class App extends XmlEntity {
    /**
     * 指定由哪个App生成或打开
     */
    private String application;
    /**
     * 是否加密，当前不支持加密
     */
    @SuppressWarnings("unused")
    private int docSecurity;
    @SuppressWarnings("unused")
    private boolean scaleCrop;
    /**
     * 公司名，可用于防伪，通过鼠标右建-&gt;详细属性查看
     */
    private String company;
    @SuppressWarnings("unused")
    private boolean linksUpToDate;
    @SuppressWarnings("unused")
    private boolean sharedDoc;
    @SuppressWarnings("unused")
    private boolean hyperlinksChanged;
    /**
     * App版本，对应{@link #application}
     */
    private String appVersion;
    /**
     * Worksheet名称集合
     */
    private TitlesOfParts titlesOfParts;
    /**
     * Worksheet数量信息
     */
    private HeadingPairs headingPairs;

    public class TitlesOfParts {
        @NS(value = "vt", contentUse = true)
        @Attr(name = {"baseType", "size"}, value = {"lpstr", "#size#"})
        List<String> vector; // sheetName

        public void setVector(final List<String> vector) {
            this.vector = vector;
            headingPairs = new HeadingPairs();
            headingPairs.vector = new ArrayList<>();
            headingPairs.vector.add(new Tuple2<>("lpstr", "Workbook"));
            headingPairs.vector.add(new Tuple2<>("i4", String.valueOf(vector.size())));
        }
    }

    private static class HeadingPairs {
        @NS(value = "vt", contentUse = true)
        @Attr(name = {"baseType", "size"}, value = {"variant", "#size#"})
        List<Tuple2<String, String>> vector;
    }

    public void setTitlePards(List<String> list) {
        if (titlesOfParts == null) {
            titlesOfParts = new TitlesOfParts();
        }
        titlesOfParts.setVector(list);
    }

    public void setApplication(String application) {
        this.application = application;
    }

    public void setDocSecurity(int docSecurity) {
        this.docSecurity = docSecurity;
    }

    public void setScaleCrop(boolean scaleCrop) {
        this.scaleCrop = scaleCrop;
    }

    public void setCompany(String company) {
        this.company = company;
    }

    public void setLinksUpToDate(boolean linksUpToDate) {
        this.linksUpToDate = linksUpToDate;
    }

    public void setSharedDoc(boolean sharedDoc) {
        this.sharedDoc = sharedDoc;
    }

    public void setHyperlinksChanged(boolean hyperlinksChanged) {
        this.hyperlinksChanged = hyperlinksChanged;
    }

    /**
     * Setting the app version, it must not be null
     *
     * @param appVersion the app version
     */
    public void setAppVersion(String appVersion) {
        if (isEmpty(appVersion)) {
            this.appVersion = "1.0.0";
        } else {
            // Filter other character but number and `.`
            char[] chars = appVersion.toCharArray();
            int i = 0, n = 0;
            for (int j = 0; j < chars.length; j++) {
                if (chars[j] >= '0' && chars[j] <= '9')
                    chars[i++] = chars[j];
                else if (chars[j] == '.' && i > 0 && chars[i - 1] != '.' && n < 2) {
                    chars[i++] = chars[j];
                    n++;
                }
                else break;
            }
            this.appVersion = i > 0 ? new String(chars, 0, chars[i - 1] != '.' ? i : i - 1) : "1.0.0";
        }
    }

    public String getApplication() {
        return application;
    }

    public String getCompany() {
        return company;
    }

    public String getAppVersion() {
        return appVersion;
    }
}

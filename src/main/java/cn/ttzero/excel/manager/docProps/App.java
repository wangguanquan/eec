/*
 * Copyright (c) 2009, guanquan.wang@yandex.com All Rights Reserved.
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
import cn.ttzero.excel.entity.NameValue;

import java.util.ArrayList;
import java.util.List;

/**
 * Created by guanquan.wang on 2017/9/21.
 */
@TopNS(prefix = {"vt", ""}, uri = {"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
        , "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"}, value = "Properties")
public class App extends XmlEntity {
    private String application;
    private int docSecurity;
    private boolean scaleCrop;
    private String manager;
    private String company;
    private boolean linksUpToDate;
    private boolean sharedDoc;
    private boolean hyperlinksChanged;
    private String appVersion;
    private TitlesOfParts titlesOfParts;
    private HeadingPairs headingPairs;

    public class TitlesOfParts {
        @NS(value = "vt", contentUse = true)
        @Attr(name = {"baseType", "size"}, value = {"lpstr", "#size#"})
        List<String> vector; // sheetName

        public void setVector(final List<String> vector) {
            this.vector = vector;
            headingPairs = new HeadingPairs();
            headingPairs.vector = new ArrayList<>();
            headingPairs.vector.add(new NameValue("lpstr", "Workbook"));
            headingPairs.vector.add(new NameValue("i4", String.valueOf(vector.size())));
        }
    }

    private static class HeadingPairs {
        @NS(value = "vt", contentUse = true)
        @Attr(name = {"baseType", "size"}, value = {"variant", "#size#"})
        List<NameValue> vector;
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

    public void setManager(String manager) {
        this.manager = manager;
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

    public void setAppVersion(String appVersion) {
        this.appVersion = appVersion;
    }

}

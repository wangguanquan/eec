package net.cua.export.manager.docProps;


import net.cua.export.annotation.Attr;
import net.cua.export.annotation.NS;
import net.cua.export.annotation.TopNS;
import net.cua.export.entity.NameValue;
import net.cua.export.entity.XmlEntity;

import java.lang.reflect.InvocationTargetException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

/**
 * Created by wanggq on 2017/9/21.
 */
@TopNS(prefix = {"vt", ""}, uri = {"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
        , "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"}, value = "Properties")
public class App extends XmlEntity {
    private String application = "Microsoft Excel";
    private int docSecurity;
    private boolean scaleCrop;
    private String manager;
    private String company;
    private boolean linksUpToDate;
    private boolean sharedDoc;
    private boolean hyperlinksChanged;
    private String appVersion = "12.0000";   // excel版本
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

    private class HeadingPairs {
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

//    public static void main(String[] args) throws IllegalAccessException, NoSuchMethodException, InvocationTargetException {
//        App app = new App();
//        app.company = "蜗牛数字有限公司";
//
//        app.titlesOfParts = app.new TitlesOfParts();
//        app.titlesOfParts.setVector(Arrays.asList("服务器列表", "测试", "IP统计", "Sheet1"));
//
//        app.writeTo("f:/app.xml");
//    }
}

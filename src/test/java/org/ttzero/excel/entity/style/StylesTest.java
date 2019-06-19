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

package org.ttzero.excel.entity.style;

import org.junit.Before;
import org.junit.Test;
import org.ttzero.excel.entity.I18N;
import org.ttzero.excel.entity.WorkbookTest;

import java.io.IOException;
import java.nio.file.Path;

import static org.ttzero.excel.entity.WorkbookTest.getOutputTestPath;
import static org.ttzero.excel.entity.style.Styles.INDEX_NUMBER_FORMAT;
import static org.ttzero.excel.entity.style.Styles.testCodeIsDate;

/**
 * Create by guanquan.wang at 2019-06-06 16:00
 */
public class StylesTest {

    private Styles styles;

    @Before public void before() {
        styles = Styles.create(new I18N());

        // Built-In number format
        styles.of(16 << INDEX_NUMBER_FORMAT);
        styles.of(20 << INDEX_NUMBER_FORMAT);
        styles.of(30 << INDEX_NUMBER_FORMAT);
        styles.of(46 << INDEX_NUMBER_FORMAT);
        styles.of(7 << INDEX_NUMBER_FORMAT); // Not data-time
        styles.of(14 << INDEX_NUMBER_FORMAT);
        styles.of(10 << INDEX_NUMBER_FORMAT); // not data-time
        styles.of(13 << INDEX_NUMBER_FORMAT); // not data-time

        // Customize
        styles.of(styles.addNumFmt(new NumFmt("\"¥\"#,##0.00;\"¥\"\\-#,##0.00")));
        styles.of(styles.addNumFmt(new NumFmt("[$-F800]dddd\\,\\ mmmm\\ dd\\,\\ yyyy")));
        styles.of(styles.addNumFmt(new NumFmt("[DBNum1][$-804]yyyy\"年\"m\"月\"d\"日\";@")));
        styles.of(styles.addNumFmt(new NumFmt("[DBNum1][$-804]yyyy\"年\"m\"月\";@")));
        styles.of(styles.addNumFmt(new NumFmt("[DBNum1][$-804]m\"月\"d\"日\";@")));
        styles.of(styles.addNumFmt(new NumFmt("yyyy\"年\"m\"月\"d\"日\";@")));
        styles.of(styles.addNumFmt(new NumFmt("[$-409]yyyy/m/d\\ h:mm\\ AM/PM;@")));
        styles.of(styles.addNumFmt(new NumFmt("yy/m/d;@")));
        styles.of(styles.addNumFmt(new NumFmt("[$-409]mmmmm/yy;@")));
        styles.of(styles.addNumFmt(new NumFmt("[$-409]d/mmm/yy;@")));
        styles.of(styles.addNumFmt(new NumFmt("[$-409]dd/mmm/yy;@")));
        styles.of(styles.addNumFmt(new NumFmt("[$-F400]h:mm:ss\\ AM/PM")));
        styles.of(styles.addNumFmt(new NumFmt("[$-409]h:mm:ss\\ AM/PM;@")));
        styles.of(styles.addNumFmt(new NumFmt("h\"时\"mm\"分\"ss\"秒\";@")));
        styles.of(styles.addNumFmt(new NumFmt("上午/下午h\"时\"mm\"分\"ss\"秒\";@")));
        styles.of(styles.addNumFmt(new NumFmt("[DBNum1][$-804]上午/下午h\"时\"mm\"分\";@")));
        styles.of(styles.addNumFmt(new NumFmt("[$-2010000]yyyy/mm/dd;@")));
        styles.of(styles.addNumFmt(new NumFmt("[$-C07]d\\.mmmm\\ yyyy;@")));
        styles.of(styles.addNumFmt(new NumFmt("[$-FC19]dd\\ mmmm\\ yyyy\\ \\г\\.;@")));
        styles.of(styles.addNumFmt(new NumFmt("[$-FC19]yyyy\\,\\ dd\\ mmmm;@")));
        styles.of(styles.addNumFmt(new NumFmt("[$-80C]dddd\\ d\\ mmmm\\ yyyy;@")));
        styles.of(styles.addNumFmt(new NumFmt("[$-44F]dd\\ mmmm\\ yyyy\\ dddd;@")));
        styles.of(styles.addNumFmt(new NumFmt("[$-816]d\\ \"de\"\\ mmmm\\ \"de\"\\ yyyy;@")));
        styles.of(styles.addNumFmt(new NumFmt("yyyy/mm/dd\\ hh:mm:ss")));
        styles.of(styles.addNumFmt(new NumFmt("yyyy/mm/dd")));
        styles.of(styles.addNumFmt(new NumFmt("m/d")));
    }

    @Test public void testTestCodeIsDate() {
        assert !testCodeIsDate("\"¥\"#,##0.00;\"¥\"\\-#,##0.00");
        assert testCodeIsDate("[$-F800]dddd\\,\\ mmmm\\ dd\\,\\ yyyy");
        assert testCodeIsDate("[DBNum1][$-804]yyyy\"年\"m\"月\"d\"日\";@");
        assert testCodeIsDate("[DBNum1][$-804]yyyy\"年\"m\"月\";@");
        assert testCodeIsDate("[DBNum1][$-804]m\"月\"d\"日\";@");
        assert testCodeIsDate("yyyy\"年\"m\"月\"d\"日\";@");
        assert testCodeIsDate("[$-409]yyyy/m/d\\ h:mm\\ AM/PM;@");
        assert testCodeIsDate("yy/m/d;@");
        assert testCodeIsDate("[$-409]mmmmm/yy;@");
        assert testCodeIsDate("[$-409]d/mmm/yy;@");
        assert testCodeIsDate("[$-409]dd/mmm/yy;@");
        assert testCodeIsDate("[$-F400]h:mm:ss\\ AM/PM");
        assert testCodeIsDate("[$-409]h:mm:ss\\ AM/PM;@");
        assert testCodeIsDate("h\"时\"mm\"分\"ss\"秒\";@");
        assert testCodeIsDate("上午/下午h\"时\"mm\"分\"ss\"秒\";@");
        assert testCodeIsDate("[DBNum1][$-804]上午/下午h\"时\"mm\"分\";@");
        assert testCodeIsDate("[$-2010000]yyyy/mm/dd;@");
        assert testCodeIsDate("[$-C07]d\\.mmmm\\ yyyy;@");
        assert testCodeIsDate("[$-FC19]dd\\ mmmm\\ yyyy\\ \\г\\.;@");
        assert testCodeIsDate("[$-FC19]yyyy\\,\\ dd\\ mmmm;@");
        assert testCodeIsDate("[$-80C]dddd\\ d\\ mmmm\\ yyyy;@");
        assert testCodeIsDate("[$-44F]dd\\ mmmm\\ yyyy\\ dddd;@");
        assert testCodeIsDate("[$-816]d\\ \"de\"\\ mmmm\\ \"de\"\\ yyyy;@");
        assert testCodeIsDate("yyyy/mm/dd\\ hh:mm:ss");
        assert testCodeIsDate("yyyy/mm/dd");
        assert testCodeIsDate("m/d");

        assert testCodeIsDate("yyyy");
        assert testCodeIsDate("m-d");
        assert testCodeIsDate("yy/m");
    }

    @Test public void testFastTestDateFmt() throws IOException {
        Path storagePath = getOutputTestPath().resolve("styles.xml");
        styles.writeTo(storagePath);

        Styles styles = Styles.load(storagePath);
        for (int i = 0, size = styles.size(); i < size; i++) {
           boolean isDate = styles.fastTestDateFmt(i);
           if (i == 0 || i == 5 || i >= 7 && i <= 9) assert !isDate;
           else assert isDate;
        }
    }
}

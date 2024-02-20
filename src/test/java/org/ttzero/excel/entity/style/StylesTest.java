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

package org.ttzero.excel.entity.style;

import org.junit.Before;
import org.junit.Test;
import org.ttzero.excel.entity.I18N;

import java.awt.Color;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertTrue;
import static org.ttzero.excel.entity.WorkbookTest.getOutputTestPath;
import static org.ttzero.excel.entity.style.Styles.INDEX_BORDER;
import static org.ttzero.excel.entity.style.Styles.INDEX_FILL;
import static org.ttzero.excel.entity.style.Styles.INDEX_FONT;
import static org.ttzero.excel.entity.style.Styles.INDEX_HORIZONTAL;
import static org.ttzero.excel.entity.style.Styles.INDEX_NUMBER_FORMAT;
import static org.ttzero.excel.entity.style.Styles.INDEX_VERTICAL;
import static org.ttzero.excel.entity.style.Styles.INDEX_WRAP_TEXT;
import static org.ttzero.excel.entity.style.Styles.testCodeIsDate;

/**
 * @author guanquan.wang at 2019-06-06 16:00
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
        assertFalse(testCodeIsDate("\"¥\"#,##0.00;\"¥\"\\-#,##0.00"));
        assertTrue(testCodeIsDate("[$-F800]dddd\\,\\ mmmm\\ dd\\,\\ yyyy"));
        assertTrue(testCodeIsDate("[DBNum1][$-804]yyyy\"年\"m\"月\"d\"日\";@"));
        assertTrue(testCodeIsDate("[DBNum1][$-804]yyyy\"年\"m\"月\";@"));
        assertTrue(testCodeIsDate("[DBNum1][$-804]m\"月\"d\"日\";@"));
        assertTrue(testCodeIsDate("yyyy\"年\"m\"月\"d\"日\";@"));
        assertTrue(testCodeIsDate("[$-409]yyyy/m/d\\ h:mm\\ AM/PM;@"));
        assertTrue(testCodeIsDate("yy/m/d;@"));
        assertTrue(testCodeIsDate("[$-409]mmmmm/yy;@"));
        assertTrue(testCodeIsDate("[$-409]d/mmm/yy;@"));
        assertTrue(testCodeIsDate("[$-409]dd/mmm/yy;@"));
        assertTrue(testCodeIsDate("[$-F400]h:mm:ss\\ AM/PM"));
        assertTrue(testCodeIsDate("[$-409]h:mm:ss\\ AM/PM;@"));
        assertTrue(testCodeIsDate("h\"时\"mm\"分\"ss\"秒\";@"));
        assertTrue(testCodeIsDate("上午/下午h\"时\"mm\"分\"ss\"秒\";@"));
        assertTrue(testCodeIsDate("[DBNum1][$-804]上午/下午h\"时\"mm\"分\";@"));
        assertTrue(testCodeIsDate("[$-2010000]yyyy/mm/dd;@"));
        assertTrue(testCodeIsDate("[$-C07]d\\.mmmm\\ yyyy;@"));
        assertTrue(testCodeIsDate("[$-FC19]dd\\ mmmm\\ yyyy\\ \\г\\.;@"));
        assertTrue(testCodeIsDate("[$-FC19]yyyy\\,\\ dd\\ mmmm;@"));
        assertTrue(testCodeIsDate("[$-80C]dddd\\ d\\ mmmm\\ yyyy;@"));
        assertTrue(testCodeIsDate("[$-44F]dd\\ mmmm\\ yyyy\\ dddd;@"));
        assertTrue(testCodeIsDate("[$-816]d\\ \"de\"\\ mmmm\\ \"de\"\\ yyyy;@"));
        assertTrue(testCodeIsDate("yyyy/mm/dd\\ hh:mm:ss"));
        assertTrue(testCodeIsDate("yyyy/mm/dd"));
        assertTrue(testCodeIsDate("m/d"));

        assertTrue(testCodeIsDate("yyyy"));
        assertTrue(testCodeIsDate("m-d"));
        assertTrue(testCodeIsDate("yy/m"));
    }

    @Test public void testFastTestDateFmt() throws IOException {
        Path storagePath = getOutputTestPath().resolve("styles.xml");
        styles.writeTo(storagePath);

        Styles styles = Styles.load(Files.newInputStream(storagePath));
        for (int i = 0, size = styles.size(); i < size; i++) {
           boolean isDate = styles.fastTestDateFmt(i);
           if (i == 0 || i == 5 || i >= 7 && i <= 9) assertFalse(isDate);
           else assertTrue(isDate);
        }
    }

    @Test public void testClear() {
        int style = (7 << INDEX_NUMBER_FORMAT) | (6 << INDEX_FONT)
                | (5 << INDEX_FILL) | (4 << INDEX_BORDER)
                | (3 << INDEX_VERTICAL) | (2 << INDEX_HORIZONTAL) | 1;

        assertEquals(Styles.clearNumFmt(style), style - (7 << INDEX_NUMBER_FORMAT));
        assertEquals(Styles.clearFont(style), style - (6 << INDEX_FONT));
        assertEquals(Styles.clearFill(style), style - (5 << INDEX_FILL));
        assertEquals(Styles.clearBorder(style), style - (4 << INDEX_BORDER));
        assertEquals(Styles.clearVertical(style), style - (3 << INDEX_VERTICAL));
        assertEquals(Styles.clearHorizontal(style), style - (2 << INDEX_HORIZONTAL));
        assertEquals(Styles.clearWrapText(style), style - (1 << INDEX_WRAP_TEXT));
    }

    @Test public void testHas() {
        int style = (7 << INDEX_NUMBER_FORMAT) | (6 << INDEX_FONT)
                | (5 << INDEX_FILL) | (4 << INDEX_BORDER)
                | (3 << INDEX_VERTICAL) | (2 << INDEX_HORIZONTAL) | 1;

        assertTrue(Styles.hasNumFmt(style));
        assertTrue(Styles.hasFont(style));
        assertTrue(Styles.hasFill(style));
        assertTrue(Styles.hasBorder(style));
        assertTrue(Styles.hasVertical(style));
        assertTrue(Styles.hasHorizontal(style));
        assertTrue(Styles.hasWrapText(style));


        assertFalse(Styles.hasNumFmt(Styles.clearNumFmt(style)));
        // Font is required
//        assertTrue(!Styles.hasFont(Styles.clearFont(style)));
        assertFalse(Styles.hasFill(Styles.clearFill(style)));
        assertFalse(Styles.hasBorder(Styles.clearBorder(style)));
        assertFalse(Styles.hasVertical(Styles.clearVertical(style)));
        assertFalse(Styles.hasHorizontal(Styles.clearHorizontal(style)));
        assertFalse(Styles.hasWrapText(Styles.clearWrapText(style)));
    }

    @Test public void testThemeColor() {
        // +-1
        Color color1 = HlsColor.calculateColor(Color.decode("#F79646"), "0.39997558519241921");
        assertTrue(color1.getRed() <= 251 && color1.getRed() >= 249);
        assertTrue(color1.getGreen() <= 192 && color1.getGreen() >= 190);
        assertTrue(color1.getBlue() <= 144 && color1.getBlue() >= 142);

        Color color2 = HlsColor.calculateColor(Color.decode("#4F81BD"), "0.79998168889431442");
        assertTrue(color2.getRed() <= 221 && color2.getRed() >= 219);
        assertTrue(color2.getGreen() <= 231 && color2.getGreen() >= 229);
        assertTrue(color2.getBlue() <= 242 && color2.getBlue() >= 240);

        Color color3 = HlsColor.calculateColor(Color.decode("#C0504D"), "0.59999389629810485");
        assertTrue(color3.getRed() <= 231 && color3.getRed() >= 229);
        assertTrue(color3.getGreen() <= 185 && color3.getGreen() >= 183);
        assertTrue(color3.getBlue() <= 184 && color3.getBlue() >= 182);

        Color color4 = HlsColor.calculateColor(new Color(0, 0, 0), "0.39997558519241921");
        assertTrue(color4.getRed() <= 103 && color4.getRed() >= 101);
        assertTrue(color4.getGreen() <= 103 && color4.getGreen() >= 101);
        assertTrue(color4.getBlue() <= 103 && color4.getBlue() >= 101);
    }

    @Test public void testRound2() {
        assertEquals(Font.round10(11), 110);
        assertEquals(Font.round10(11.1), 110);
        assertEquals(Font.round10(11.2), 110);
        assertEquals(Font.round10(11.22), 110);
        assertEquals(Font.round10(11.23), 115);
        assertEquals(Font.round10(11.3), 115);
        assertEquals(Font.round10(11.5), 115);
        assertEquals(Font.round10(11.7), 115);
        assertEquals(Font.round10(11.72), 115);
        assertEquals(Font.round10(11.73), 120);
        assertEquals(Font.round10(11.8), 120);

        Font font = Font.parse("italic_bold_12.24_宋体");
        assertTrue(font.isItalic());
        assertTrue(font.isBold());
        assertTrue(font.getSize2() - 12.5D <= 0.0001);
        assertEquals(font.getName(), "宋体");
    }

    @Test public void testFontConversion() {
        // awt Font to eec Font
        java.awt.Font awtFont = new java.awt.Font("Arial", java.awt.Font.PLAIN, 12);
        Font font = Font.of(awtFont);

        assertEquals(font.getName(), awtFont.getName());
        assertEquals(font.getSize(), awtFont.getSize());
        assertEquals(font.getStyle(), Font.Style.PLAIN);

        // eec Font to awt Font
        java.awt.Font awtFont2 = font.toAwtFont();
        assertEquals(awtFont, awtFont2);

        awtFont = new java.awt.Font("宋体", java.awt.Font.BOLD, 16);
        font = Font.of(awtFont);
        assertEquals(font.getName(), awtFont.getName());
        assertEquals(font.getSize(), awtFont.getSize());
        assertTrue(font.isBold());

        awtFont2 = font.toAwtFont();
        assertEquals(awtFont, awtFont2);

        awtFont = new java.awt.Font("宋体", java.awt.Font.BOLD | java.awt.Font.ITALIC, 16);
        font = Font.of(awtFont);
        assertEquals(font.getName(), awtFont.getName());
        assertEquals(font.getSize(), awtFont.getSize());
        assertTrue(font.isBold() && font.isItalic());

        awtFont2 = font.toAwtFont();
        assertEquals(awtFont, awtFont2);
    }
}

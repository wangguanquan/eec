/*
 * Copyright (c) 2017-2023, guanquan.wang@yandex.com All Rights Reserved.
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

import org.dom4j.Document;
import org.dom4j.Element;
import org.dom4j.io.SAXReader;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.ttzero.excel.manager.TopNS;

import java.awt.Color;
import java.io.InputStream;
import java.util.List;

import static org.ttzero.excel.entity.style.Styles.getAttr;
import static org.ttzero.excel.util.ExtBufferedWriter.digits_uppercase;

/**
 * Theme style
 * NOTE: Only read the theme color current
 *
 * @author guanquan.wang at 2023-01-18 10:29
 */
@TopNS(prefix = "", uri = "http://schemas.openxmlformats.org/drawingml/2006/main", value = "a")
public class Theme {
    /**
     * LOGGER
     */
    private static final Logger LOGGER = LoggerFactory.getLogger(Theme.class);
    // theme colors
    private ClrScheme[] clrs;

    public static Theme load(InputStream is) {
        Theme self = new Theme();
        // load theme1.xml
        SAXReader reader = SAXReader.createDefault();
        Document document;
        try {
            document = reader.read(is);
            Element root = document.getRootElement();
            Element clrScheme = root.element("themeElements").element("clrScheme");
            List<Element> clrSchemes;
            if (clrScheme != null && (clrSchemes = clrScheme.elements()) != null && !clrSchemes.isEmpty()) {
                ClrScheme[] clrs = new ClrScheme[clrSchemes.size()];
                self.clrs = clrs;
                int i = 0;
                for (Element e : clrSchemes) self.clrs[i++] = toClrScheme(e);

                // Adjust color index (lt1 > dk1 > lt2 > dk2)
                if (clrs.length >= 2 && "dk1".equals(clrs[0].tag) && "lt1".equals(clrs[1].tag)) {
                    ClrScheme tmp = clrs[1];
                    clrs[1] = clrs[0];
                    clrs[0] = tmp;
                }
                if (clrs.length >= 4 && "dk2".equals(clrs[2].tag) && "lt2".equals(clrs[3].tag)) {
                    ClrScheme tmp = clrs[3];
                    clrs[3] = clrs[2];
                    clrs[2] = tmp;
                }

                // FIXME Temporary processing
                int len = Math.min(clrs.length, ColorIndex.themeColors.length);
                i = 0;
                for (; i < len; i++) ColorIndex.themeColors[i] = clrs[i].color;
            }

            // TODO others

        } catch (Exception e) {
            LOGGER.warn("Read the theme failed and ignore the style to continue.", e);
            // Ignore
        }
        return self;
    }

    static ClrScheme toClrScheme(Element e) {
        ClrScheme c = new ClrScheme();
        c.tag = e.getName();
        List<Element> subs = e.elements();
        Color color = null;
        if (subs != null && !subs.isEmpty()) {
            Element sub = subs.get(0);
            String v = getAttr(sub, "lastClr");
            if (v == null) v = getAttr(sub, "val");
            else {
                try {
                    color = ColorIndex.toColor(v);
                } catch (Exception ex) {
                    v = getAttr(sub, "val");
                }
            }
            if (color == null) {
                try {
                    color = ColorIndex.toColor(v);
                } catch (Exception ex) {
                    color = new BuildInColor(64); // auto if exception
                }
            }
        } else color = new BuildInColor(64); // auto if unknown tag
        c.color = color;
        return c;
    }

    public Color[] getClrSchemes() {
        if (clrs != null && clrs.length > 0) {
            Color[] colors = new Color[clrs.length];
            for (int i = 0; i < clrs.length; i++) colors[i] = clrs[i].color;
            return colors;
        } else return ColorIndex.themeColors;
    }

    public static class ClrScheme {
        public String tag;
        public Color color;

        public ClrScheme() { }

        public ClrScheme(String tag, Color color) {
            this.tag = tag;
            this.color = color;
        }

        public String toString() {
            if (color == null) return tag;
            int r = color.getRed(), g = color.getGreen(), b = color.getBlue();
            return tag + ": " + new String(new char[] {
                digits_uppercase[r >> 4], digits_uppercase[r & 0xF],
                digits_uppercase[g >> 4], digits_uppercase[g & 0xF],
                digits_uppercase[b >> 4], digits_uppercase[b & 0xF]
            });
        }
    }
}

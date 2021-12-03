/*
 * Copyright (c) 2017-2020, guanquan.wang@yandex.com All Rights Reserved.
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


package org.ttzero.excel.entity;

import org.ttzero.excel.annotation.TopNS;
import org.ttzero.excel.entity.style.Font;
import org.ttzero.excel.manager.Const;
import org.ttzero.excel.reader.ExcelReader;
import org.ttzero.excel.util.ExtBufferedWriter;
import org.ttzero.excel.util.FileUtil;

import java.io.Closeable;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;

import static org.ttzero.excel.util.StringUtil.isNotEmpty;

/**
 * A simple implementation
 *
 * @author guanquan.wang at 2020-05-21 16:31
 */
@TopNS(prefix = "", value = "comments", uri = Const.SCHEMA_MAIN)
public class Comments implements Storable, Closeable {

    /** Comments Cache*/
    private final List<C> commentList;
    private final int id;
    private final String author;
//    private final static int CACHE_SIZE = 20;

    Comments(int id, String author) {
        this.id = id;
        this.author = author;
        commentList = new ArrayList<>();
    }

    public void addComment(String ref, String title, String value) {
        C c = new C();
        c.ref = ref;
        c.text = new ArrayList<>();
        if (isNotEmpty(title)) {
            parse(title, true, c.text);
        }
        if (isNotEmpty(value)) {
            parse(value, false, c.text);
        }

        commentList.add(c);

//        if (commentList.size() >= CACHE_SIZE) {
//            flush();
//        }
    }

    private void parse(String val, boolean bold, List<R> list) {
        // a simple implement
        R r = new R();
        r.rPr = new Pr("宋体", 9);
        if (bold) {
            r.rPr.bold();
            if (val.charAt(val.length() - 1) != 10) {
                val += (char) 10;
            }
        }
        r.t = val;
        list.add(r);
    }

    public void flush() {
        // TODO Write tmp and clear cache

    }

    @Override
    public void close() {
        // Ignore
    }

    @Override
    public void writeTo(Path root) throws IOException {
        if (commentList.isEmpty()) return;
        try (ExtBufferedWriter writer = new ExtBufferedWriter(
            Files.newBufferedWriter(root.resolve("comments" + id + Const.Suffix.XML)))) {
            writer.write(Const.EXCEL_XML_DECLARATION);
            writer.newLine();
            TopNS topNS = this.getClass().getAnnotation(TopNS.class);
            writer.write('<');
            writer.write(topNS.value());
            writer.write(" xmlns=\"");
            writer.write(topNS.uri()[0]);
            writer.write("\"><authors><author>");
            writer.escapeWrite(isNotEmpty(author) ? author : System.getProperty("user.name"));
            writer.write("</author></authors><commentList>");

            for (C c : commentList) {
                writer.write("<comment ref=\"");
                writer.write(c.ref);
                writer.write("\" authorId=\"0\"><text>");
                for (R r : c.text) {
                    writer.write("<r>");
                    writer.write(r.rPr.toString());
                    writer.write("<t");
                    writer.write((r.t.indexOf(10) > 0 ? " xml:space=\"preserve\">" : ">"));
                    writer.escapeWrite(r.t);
                    writer.write("</t></r>");
                }
                writer.write("</text></comment>");
            }
            writer.write("</commentList></comments>");
        }

        // Write vml
        vml(root);
    }

    private void vml(Path root) throws IOException {
        Path parent = root.resolve("drawings");
        if (!Files.exists(parent)) {
            FileUtil.mkdir(parent);
        }

        try (ExtBufferedWriter writer = new ExtBufferedWriter(
            Files.newBufferedWriter(parent.resolve("vmlDrawing" + id + Const.Suffix.VML)))) {
            writer.write("<xml xmlns:v=\"urn:schemas-microsoft-com:vml\"");
            writer.write(" xmlns:o=\"urn:schemas-microsoft-com:office:office\"");
            writer.write(" xmlns:x=\"urn:schemas-microsoft-com:office:excel\">");
            writer.write(" <o:shapelayout v:ext=\"edit\">");
            writer.write("  <o:idmap v:ext=\"edit\" data=\"1\"/>");
            writer.write(" </o:shapelayout>");
            writer.write(" <v:shapetype id=\"_x0000_t202\" coordsize=\"21600,21600\" o:spt=\"202\"");
            writer.write("  path=\"m,l,21600r21600,l21600,xe\">");
            writer.write("  <v:stroke joinstyle=\"miter\"/>");
            writer.write("  <v:path gradientshapeok=\"t\" o:connecttype=\"rect\"/>");
            writer.write(" </v:shapetype>");
            int i = 1;
            for (C c : commentList) {
                long cr = ExcelReader.cellRangeToLong(c.ref);
                writer.write(" <v:shape id=\"_x0000_s");writer.write(100 + i);
                writer.write("\" type=\"#_x0000_t202\" style='width:100.8pt;height:60.6pt;z-index:");
                writer.writeInt(i++);
                writer.write(";  visibility:hidden' fillcolor=\"#ffffe1\" o:insetmode=\"auto\">");
                writer.write("  <v:fill color2=\"#ffffe1\"/>");
                writer.write("  <v:shadow on=\"t\" color=\"black\" obscured=\"t\"/>");
                writer.write("  <v:path o:connecttype=\"none\"/>");
                writer.write("  <v:textbox style='mso-direction-alt:auto'>");
                writer.write("   <div style='text-align:left'></div>");
                writer.write("  </v:textbox>");
                writer.write("  <x:ClientData ObjectType=\"Note\">");
                writer.write("   <x:MoveWithCells/>");
                writer.write("   <x:SizeWithCells/>");
                writer.write("   <x:Anchor/>");
                writer.write("   <x:AutoFill>False</x:AutoFill>");
                writer.write("   <x:Row>");writer.write((cr >> 16) - 1);writer.write("</x:Row>");
                writer.write("   <x:Column>");writer.write((cr & 0x7FFF) - 1);writer.write("</x:Column>");
                writer.write("  </x:ClientData>");
                writer.write(" </v:shape>");
            }
            writer.write("</xml>");
        }
    }

    private static class C {
        private String ref;
        private List<R> text;

        @Override
        public String toString() {
            StringBuilder buf = new StringBuilder("<comment ref=\"")
                .append(ref).append("\" authorId=\"0\"><text>");
            for (R r : text)
                buf.append(r);
            buf.append("</text>").append("</comment>");
            return buf.toString();
        }
    }

    private static class R {
        private Pr rPr;
        private String t;

        @Override
        public String toString() {
            return "<r>" + rPr + "<t" + (t.indexOf(10) > 0 ? " xml:space=\"preserve\">" : ">") +
                t + "</t>" + "</r>";
        }
    }

    private static class Pr extends Font {
        private static final String[] STYLE = {"", "<u/>", "<b/>", "<u/><b/>", "<i/>", "<i/><u/>", "<b/><i/>", "<i/><b/><u/>"};
        Pr(String name, int size) {
            super(name, size);
        }

        @Override
        public String toString() {
            return "<rPr>" + STYLE[getStyle() & 0x07] +
                "<sz val=\"" + getSize() + "\"/>" +
                "<rFont val=\"" + getName() + "\"/>" +
                "<charset val=\"" + getCharset() + "\"/></rPr>";
        }
    }
}

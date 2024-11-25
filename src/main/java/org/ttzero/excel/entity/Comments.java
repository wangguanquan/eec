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

import org.ttzero.excel.entity.style.ColorIndex;
import org.ttzero.excel.manager.TopNS;
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

import static org.ttzero.excel.entity.Sheet.toCoordinate;
import static org.ttzero.excel.util.StringUtil.isNotEmpty;

/**
 * 批注的简单实现
 *
 * @author guanquan.wang at 2020-05-21 16:31
 */
@TopNS(prefix = "", value = "comments", uri = Const.SCHEMA_MAIN)
public class Comments implements Storable, Closeable {
    /**
     * 换行符常量，用于在添加批注时分隔标题和内容
     */
    private static final String LF = "\n";

    /** Comments Cache*/
    public List<C> commentList = new ArrayList<>();
    public int id;
    public String author;

    public Comments() { }

    public Comments(int id, String author) {
        this.id = id;
        this.author = author;
    }

    /**
     * 添加评论到指定的单元格位置
     *
     * @param ref 单元格位置索引
     * @param title 评论的标题
     * @param value 评论的内容
     * @return 包含新添加批注的对象
     */
    public C addComment(String ref, String title, String value) {
        return addComment(ref, new Comment(title, value));
    }

    /**
     * 添加评论到指定的单元格位置
     *
     * @param ref 单元格位置索引
     * @param comment 批注对象
     * @return 包含新添加批注的对象
     */
    public C addComment(String ref, Comment comment) {
        C c = new C();
        c.ref = ref;
        c.width = comment.getWidth();
        c.height = comment.getHeight();
        boolean hasTitle = isNotEmpty(comment.getTitle()), hasValue = isNotEmpty(comment.getValue());
        c.nodes = new R[hasTitle && hasValue ? 2 : 1];
        int i = 0;
        if (hasTitle) c.nodes[i++] = toR(comment.getTitle(), true, comment.getTitleFont());
        if (hasValue) c.nodes[i] = toR(comment.getValue(), false, comment.getValueFont());
        commentList.add(c);
        return c;
    }

    /**
     * 在指定行列添加批注
     *
     * @param row 行号
     * @param col 列号
     * @param value 批注内容
     * @return 包含新添加批注的对象
     */
    public C addComment(int row, int col, String value) {
        return addComment(toCoordinate(row, col), new Comment(null, value));
    }

    /**
     * 在指定单元格添加批注
     *
     * @param row 行号
     * @param col 列号
     * @param title 批注标题
     * @param value 批注内容
     * @return 包含新添加批注的对象
     */
    public C addComment(int row, int col, String title, String value) {
        return addComment(toCoordinate(row, col), new Comment(title, value));
    }

    /**
     * 在指定单元格添加批注
     *
     * @param row 行号
     * @param col 列号
     * @param comment 批注对象
     * @return 包含新添加批注的对象
     */
    public C addComment(int row, int col, Comment comment) {
        return addComment(toCoordinate(row, col), comment);
    }

    protected R toR(String val, boolean isTitle,  Font font) {
        // a simple implement
        R r = new R();
        r.rPr = font == null ? new Pr("宋体", 9) : new Pr(font);
        // set bold default
        if (isTitle) r.rPr.bold();
        r.t = val;
        return r;
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
                R[] nodes = c.nodes;
                if (c.nodes.length == 2) {
                    R lf = new R();
                    lf.t = LF;
                    lf.rPr = c.nodes[0].rPr;
                    nodes = new R[] { c.nodes[0], lf, c.nodes[1] };
                }
                for (R r : nodes) {
                    writer.write("<r>");
                    writer.write(r.rPr.toString());
                    writer.write("<t");
                    writer.write((r.t.indexOf(10) >= 0 ? " xml:space=\"preserve\">" : ">"));
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

    protected void vml(Path root) throws IOException {
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
                long cr = ExcelReader.coordinateToLong(c.ref);
                writer.write(" <v:shape id=\"_x0000_s");writer.write(100 + i);
                writer.write("\" type=\"#_x0000_t202\" style='width:" + (c.width != null ? c.width : 100.8D) + "pt;height:" + (c.height != null ? c.height : 60.6D) + " pt;z-index:");
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

    public static class C {
        public String ref;
        public R[] nodes;
        public Double width, height;

        @Override
        public String toString() {
            StringBuilder buf = new StringBuilder("<comment ref=\"")
                .append(ref).append("\" authorId=\"0\"><text>");
            for (R r : nodes)
                buf.append(r);
            buf.append("</text>").append("</comment>");
            return buf.toString();
        }
    }

    public static class R {
        public Pr rPr;
        public String t;

        @Override
        public String toString() {
            return "<r>" + rPr + "<t" + (t.indexOf(10) > 0 ? " xml:space=\"preserve\">" : ">") +
                t + "</t>" + "</r>";
        }
    }

    public static class Pr extends Font {
        public static final String[] STYLE = {"", "<u/>", "<b/>", "<u/><b/>", "<i/>", "<i/><u/>", "<b/><i/>", "<i/><b/><u/>"};
        public Pr(String name, int size) {
            super(name, size);
        }

        public Pr(Font font) {
            super(font);
        }

        @Override
        public String toString() {
            StringBuilder buf = new StringBuilder("<rPr>");
            if (getStyle() > 0 && getStyle() < 8) buf.append(STYLE[getStyle() & 0x07]);
            buf.append("<rFont val=\"").append(getName()).append("\"/>");
            buf.append("<sz val=\"").append(getSize()).append("\"/>");
            if (getCharset() > 0) buf.append("<charset val=\"").append(getCharset()).append("\"/>");
            if (getColor() != null) buf.append("<color rgb=\"").append(ColorIndex.toARGB(getColor().getRGB())).append("\"/>");
            if (getFamily() > 0) buf.append("<family val=\"").append(getFamily()).append("\"/>");
            buf.append("</rPr>");
            return buf.toString();
        }
    }
}

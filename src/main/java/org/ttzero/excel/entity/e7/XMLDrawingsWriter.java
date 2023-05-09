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


package org.ttzero.excel.entity.e7;

import org.ttzero.excel.entity.IDrawingsWriter;
import org.ttzero.excel.entity.Picture;
import org.ttzero.excel.entity.Relationship;
import org.ttzero.excel.manager.Const;
import org.ttzero.excel.manager.RelManager;
import org.ttzero.excel.manager.TopNS;
import org.ttzero.excel.util.ExtBufferedWriter;
import org.ttzero.excel.util.FileUtil;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;

import static org.ttzero.excel.util.FileUtil.exists;

/**
 * Drawings writer(For Picture only)
 *
 * @author wangguanquan3 at 2023-03-07 09:09
 */
@TopNS(prefix = {"xdr", "a", "r"}, value = "wsDr"
    , uri = {"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
    , "http://schemas.openxmlformats.org/drawingml/2006/main"
    , Const.Relationship.RELATIONSHIP})
public class XMLDrawingsWriter implements IDrawingsWriter {
    protected Path path;
    protected ExtBufferedWriter bw;
    protected int size;
    protected RelManager relManager;

    public XMLDrawingsWriter() { }

    // FIXME 临时代码
    public XMLDrawingsWriter(Path path) {
        this.path = path;
        this.relManager = new RelManager();
        try {
            if (!exists(path.getParent())) {
                FileUtil.mkdir(path.getParent());
            }
            bw = new ExtBufferedWriter(Files.newBufferedWriter(path, StandardCharsets.UTF_8));
            bw.write("<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">");
        } catch (IOException e) {
            throw new RuntimeException("Create XMLDrawingsWriter error", e);
        }
    }

    @Override
    public void close() throws IOException {
        // End tag
        if (size > 0) bw.write("</xdr:wsDr>");
        relManager.write(path.getParent(), path.getFileName().toString());
        FileUtil.close(bw);
    }

    @Override
    public void writeTo(Path root) throws IOException {
//        DocumentFactory factory = DocumentFactory.getInstance();
//        TopNS ns = XMLDrawingsWriter.class.getAnnotation(TopNS.class);
//        Namespace xdr = Namespace.get(ns.prefix()[0], ns.uri()[0])
//            , a = Namespace.get(ns.prefix()[1], ns.uri()[1])
//            , r = Namespace.get(ns.prefix()[2], ns.uri()[2]);
//        Element rootElement = factory.createElement(QName.get(ns.value(), xdr));
//        rootElement.add(a);
//        rootElement.add(r);
//
//        for (Picture p : pictures) {
//            Element anchor = rootElement.addElement(QName.get("twoCellAnchor", xdr));
//            // If Not MOVE_AND_RESIZE
//            Element from = anchor.addElement(QName.get("from", xdr)), to = anchor.addElement(QName.get("to", xdr));
//        }
//
//        FileUtil.writeToDiskNoFormat(factory.createDocument(rootElement), root);
    }

    static final String[] ANCHOR_PROPERTY = {"twoCell", "oneCell", "absolute"};

    @Override
    public void add(Picture picture) throws IOException {
//        pictures.add(picture);
        Relationship picRel = relManager.add(new Relationship("../media/" + picture.picName, Const.Relationship.IMAGE));
        size++;

        bw.write("<xdr:twoCellAnchor editAs=\"");
        bw.write(ANCHOR_PROPERTY[picture.property & 3]);
        // Default editAs="twoCell"
        bw.write("\">");

        // twoCell to 可以from在x和y+1
        // oneCell to 可以和from的x,y一致，根据size计算colOff和rowOff
        // absolute 需要通过size计算

        // From
        bw.write("<xdr:from><xdr:col>");
        bw.writeInt(picture.col);
        bw.write("</xdr:col><xdr:colOff>");
        bw.writeInt((picture.padding >>> 24) * 12700);
        bw.write("</xdr:colOff><xdr:row>");
        bw.writeInt(picture.row - 1);
        bw.write("</xdr:row><xdr:rowOff>");
        bw.writeInt(((picture.padding >>> 16) & 0xFF) * 12700);
        bw.write("</xdr:rowOff></xdr:from>");

        // TO
        bw.write("<xdr:to><xdr:col>");
        bw.writeInt(picture.col + 1);
        bw.write("</xdr:col><xdr:colOff>");
        bw.writeInt(-((picture.padding >>> 8) & 0xFF) * 12700);
        bw.write("</xdr:colOff><xdr:row>");
        bw.writeInt(picture.row);
        bw.write("</xdr:row><xdr:rowOff>");
        bw.writeInt(-(picture.padding & 0xFF) * 12700);
        bw.write("</xdr:rowOff></xdr:to>");

        // Picture
        bw.write("<xdr:pic><xdr:nvPicPr><xdr:cNvPr id=\"");
        bw.writeInt(size);
        bw.write("\" name=\"Picture ");
        bw.writeInt(size);
        bw.write("\"/><xdr:cNvPicPr><a:picLocks noChangeAspect=\"1\"/></xdr:cNvPicPr></xdr:nvPicPr>");
        bw.write("<xdr:blipFill><a:blip r:embed=\"");
        bw.write(picRel.getId());
        bw.write("\"/>");

        bw.write("<a:stretch><a:fillRect/></a:stretch></xdr:blipFill>");
        bw.write("<xdr:spPr><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></xdr:spPr>");
//        bw.write("<xdr:spPr><a:prstGeom prst=\"roundRect\"><a:avLst><a:gd name=\"adj\" fmla=\"val 20594\"/></a:avLst></a:prstGeom></xdr:spPr>");
        // End Picture
        bw.write("</xdr:pic><xdr:clientData/></xdr:twoCellAnchor>");
    }

}

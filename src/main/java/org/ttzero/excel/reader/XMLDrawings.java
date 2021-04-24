/*
 * Copyright (c) 2017-2021, guanquan.wang@yandex.com All Rights Reserved.
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

import org.dom4j.Document;
import org.dom4j.DocumentException;
import org.dom4j.Element;
import org.dom4j.Namespace;
import org.dom4j.QName;
import org.dom4j.io.SAXReader;
import org.ttzero.excel.entity.Relationship;
import org.ttzero.excel.manager.Const;
import org.ttzero.excel.manager.RelManager;
import org.ttzero.excel.util.FileUtil;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Collectors;

/**
 * Drawings resources
 *
 * @author guanquan.wang at 2021-04-24 16:18
 */
public class XMLDrawings implements Drawings {
    /**
     * Root path
     */
    private ExcelReader excelReader;
    /**
     * Source
     */
    private List<Drawings.Picture> pictures;
    /**
     * Mark parse flag
     */
    private boolean parsed;

    XMLDrawings(ExcelReader reader) {
        this.excelReader = reader;
    }

    /**
     * List all picture in excel
     *
     * @return list of {@link Drawings.Picture}, or null if not exists.
     */
    @Override
    public List<Drawings.Picture> listPictures() {
        return parsed ? pictures : parse();
    }

    /**
     * List all picture in specify worksheet
     *
     * @return list of {@link Drawings.Picture}, or null if not exists.
     */
    @Override
    public List<Drawings.Picture> listPictures(Sheet sheet) {
        if (!parsed) parse();
        if (pictures != null)
            return pictures.stream().filter(p -> p.sheet.getIndex() == sheet.getIndex()).collect(Collectors.toList());
        return null;
    }

    private List<Drawings.Picture> parse() {
        parsed = true;
        // Empty excel, maybe throw exception here
        if (excelReader.sheets == null) return null;

        SAXReader reader = new SAXReader();
        Document document;

        List<Picture> pictures = new ArrayList<>();
        for (Sheet sheet : excelReader.sheets) {
            XMLSheet xmlSheet = (XMLSheet) sheet;
            Path relsPath = xmlSheet.path.getParent().resolve("_rels/" + xmlSheet.path.getFileName() + ".rels");
            if (!Files.exists(relsPath)) continue;
            try {
                document = reader.read(Files.newInputStream(relsPath));
            } catch (DocumentException | IOException e) {
                FileUtil.rm_rf(excelReader.self.toFile(), true);
                throw new ExcelReadException("The file format is incorrect or corrupted. [/xl/worksheets/_rels/" + xmlSheet.path.getFileName() + ".rels]");
            }

            List<Element> list = document.getRootElement().elements();
            for (Element e : list) {
                String target = e.attributeValue("Target"), type = e.attributeValue("Type");
                // Background
                if (Const.Relationship.IMAGE.equals(type)) {
                    Picture picture = new Picture();
                    pictures.add(picture);
                    picture.sheet = sheet;
                    picture.background = true;
                    picture.localPath = xmlSheet.path.getParent().resolve(target);
                    // Drawings
                } else if (Const.Relationship.DRAWINGS.equals(type)) {
                    Path drawingsPath = xmlSheet.path.getParent().resolve(target);
                    List<Picture> subPictures = parseDrawings(drawingsPath);
                    if (subPictures != null) {
                        for (Picture picture : subPictures) {
                            picture.sheet = sheet;
                            pictures.add(picture);
                        }
                    }
                }
            }
        }

        return !pictures.isEmpty() ? (this.pictures = pictures) : null;
    }

    // Parse drawings.xml
    private List<Picture> parseDrawings(Path path) {
        SAXReader reader = new SAXReader();
        Document document;
        try {
            document = reader.read(Files.newInputStream(path.getParent().resolve("_rels/" + path.getFileName() + ".rels")));
        } catch (DocumentException | IOException e) {
            FileUtil.rm_rf(excelReader.self.toFile(), true);
            throw new ExcelReadException("The file format is incorrect or corrupted. [/xl/drawings/_rels/" + path.getFileName() + ".rels]");
        }
        List<Element> list = document.getRootElement().elements();
        Relationship[] rels = new Relationship[list.size()];
        int i = 0;
        for (Element e : list) {
            rels[i++] = new Relationship(e.attributeValue("Id"), e.attributeValue("Target"), e.attributeValue("Type"));
        }
        RelManager relManager = RelManager.of(rels);

        try {
            document = reader.read(Files.newInputStream(path));
        } catch (DocumentException | IOException e) {
            FileUtil.rm_rf(excelReader.self.toFile(), true);
            throw new ExcelReadException("The file format is incorrect or corrupted. [/xl/drawings/" + path.getFileName() + "]");
        }

        Element root = document.getRootElement();
        Namespace xdr = root.getNamespaceForPrefix("xdr"), a = root.getNamespaceForPrefix("a");

        List<Element> elements = root.elements();
        List<Picture> pictures = new ArrayList<>(elements.size());
        for (Element e : root.elements()) {
            Element pic = e.element(QName.get("pic", xdr));
            // Not a picture
            if (pic == null) continue;

            Element blipFill = pic.element(QName.get("blipFill", xdr));
            if (blipFill == null) continue;

            Element blip = blipFill.element(QName.get("blip", a));
            if (blip == null) continue;

            Namespace r = blip.getNamespaceForPrefix("r");
            String embed = blip.attributeValue(QName.get("embed", r));
            Relationship rel = relManager.getById(embed);
            if (r != null && Const.Relationship.IMAGE.equals(rel.getType())) {
                Picture picture = new Picture();
                pictures.add(picture);
                picture.localPath = path.getParent().resolve(rel.getTarget());
                picture.dimension = dimension(e, xdr);

                Element extLst = blip.element(QName.get("extLst", a));
                if (extLst == null) continue;

                for (Element ext : extLst.elements()) {
                    Element srcUrl = ext.element("picAttrSrcUrl");
                    // hyperlink
                    if (srcUrl != null) {
                        rel = relManager.getById(srcUrl.attributeValue(QName.get("id", r)));
                        if (rel != null && Const.Relationship.HYPERLINK.equals(rel.getType())) {
                            picture.srcUrl = rel.getTarget();
                        }
                    }
                }
            }
        }
        return !pictures.isEmpty() ? pictures : null;
    }

    Dimension dimension(Element e, Namespace xdr) {
        Element fromEle = e.element(QName.get("from", xdr));
        int[] f = dimEle(fromEle, xdr);
        Element toEle = e.element(QName.get("to", xdr));
        int[] t = dimEle(toEle, xdr);

        return new Dimension(f[0] + 1, (short) (f[1] + 1), t[0] + 1, (short) (t[1] + 1));
    }

    private int[] dimEle(Element e, Namespace xdr) {
        int c = 0, r = 0;
        if (e != null) {
            String col = e.element(QName.get("col", xdr)).getText(), row = e.element(QName.get("row", xdr)).getText();
            c = Integer.parseInt(col);
            r = Integer.parseInt(row);
        }
        return new int[] { r, c };
    }
}

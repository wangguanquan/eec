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
import java.nio.file.StandardCopyOption;
import java.util.ArrayList;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;

import static org.ttzero.excel.reader.ExcelReader.getEntry;
import static org.ttzero.excel.reader.ExcelReader.toZipPath;

/**
 * Drawings resources
 *
 * @author guanquan.wang at 2021-04-24 16:18
 */
public class XMLDrawings implements Drawings {
    /**
     * Root path
     */
    private final ExcelReader excelReader;
    /**
     * Source
     */
    private List<Drawings.Picture> pictures;
    /**
     * Mark parse flag
     */
    private boolean parsed;

    public XMLDrawings(ExcelReader reader) {
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
     * Parse picture
     *
     * @return list of {@link Drawings.Picture}
     */
    protected List<Drawings.Picture> parse() {
        parsed = true;
        // Empty excel, maybe throw exception here
        if (excelReader.sheets == null) return null;

        ZipFile zipFile = excelReader.zipFile;
        if (zipFile == null) return null;

        SAXReader reader = SAXReader.createDefault();
        Document document;

        List<Picture> pictures = new ArrayList<>();
        for (Sheet sheet : excelReader.sheets) {
            XMLSheet xmlSheet = (XMLSheet) sheet;
            int i = xmlSheet.path.lastIndexOf('/');
            if (i < 0) i = xmlSheet.path.lastIndexOf('\\');
            String fileName = xmlSheet.path.substring(i + 1);
            ZipEntry entry = getEntry(zipFile, "xl/worksheets/_rels/" + fileName + ".rels");
            if (entry == null) continue;
            try {
                document = reader.read(zipFile.getInputStream(entry));
            } catch (DocumentException | IOException e) {
                throw new ExcelReadException("The file format is incorrect or corrupted. [" + entry.getName() + ".rels]");
            }

            if (excelReader.tempDir == null) {
                try {
                    excelReader.tempDir = FileUtil.mktmp("eec-");
                } catch (IOException e) {
                    throw new ExcelReadException("Create temp directory failed.", e);
                }
            }
            Path imagesPath = excelReader.tempDir.resolve("media");
            if (!Files.exists(imagesPath)) {
                // Create media path
                try {
                    Files.createDirectory(imagesPath);
                } catch (IOException e) {
                    throw new ExcelReadException("Create temp directory failed.", e);
                }
            }
            List<Element> list = document.getRootElement().elements();
            for (Element e : list) {
                String target = e.attributeValue("Target"), type = e.attributeValue("Type");
                entry = getEntry(zipFile, "xl/" + toZipPath(target));
                // Background
                if (Const.Relationship.IMAGE.equals(type)) {
                    Picture picture = new Picture();
                    pictures.add(picture);
                    picture.sheet = sheet;
                    picture.background = true;
                    // Copy image to tmp file
                    try {
                        Path targetPath = imagesPath.resolve(target);
                        Files.copy(zipFile.getInputStream(entry), targetPath, StandardCopyOption.REPLACE_EXISTING);
                        picture.localPath = targetPath;
                    } catch (IOException ioException) {
                        ioException.printStackTrace();
                    }

                    // Drawings
                } else if (Const.Relationship.DRAWINGS.equals(type)) {
                    List<Picture> subPictures = parseDrawings(zipFile, entry, imagesPath);
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
    protected List<Picture> parseDrawings(ZipFile zipFile, ZipEntry entry, Path imagesPath) {
        int i = entry.getName().lastIndexOf('/');
        String relsKey;
        if (i > 0)
            relsKey = entry.getName().substring(0, i) + "/_rels" + entry.getName().substring(i);
        else if ((i = entry.getName().lastIndexOf('\\')) > 0)
            relsKey = entry.getName().substring(0, i) + "\\_rels" + entry.getName().substring(i);
        else relsKey = entry.getName();
        String key = relsKey + ".rels";
        ZipEntry entry1 = getEntry(zipFile, key);
        if (entry1 == null) throw new ExcelReadException("The file format is incorrect or corrupted. [" + key + "]");
        SAXReader reader = SAXReader.createDefault();
        Document document;
        try {
            document = reader.read(zipFile.getInputStream(entry1));
        } catch (DocumentException | IOException e) {
            throw new ExcelReadException("The file format is incorrect or corrupted. [" + key + "]");
        }
        List<Element> list = document.getRootElement().elements();
        Relationship[] rels = new Relationship[list.size()];
        i = 0;
        for (Element e : list) {
            rels[i++] = new Relationship(e.attributeValue("Id"), e.attributeValue("Target"), e.attributeValue("Type"));
        }
        RelManager relManager = RelManager.of(rels);

        try {
            document = reader.read(zipFile.getInputStream(entry));
        } catch (DocumentException | IOException e) {
            throw new ExcelReadException("The file format is incorrect or corrupted. [" + entry.getName() + "]");
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
            if (rel != null && Const.Relationship.IMAGE.equals(rel.getType())) {
                Picture picture = new Picture();
                pictures.add(picture);
                // Copy image to tmp path
                entry = getEntry(zipFile, "xl/" + toZipPath(rel.getTarget()));
                if (entry != null) {
                    try {
                        Path targetPath = imagesPath.resolve(rel.getTarget());
                        Files.copy(zipFile.getInputStream(entry), targetPath, StandardCopyOption.REPLACE_EXISTING);
                        picture.localPath = targetPath;
                    } catch (IOException ioException) {
                        ioException.printStackTrace();
                    }
                }
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

    protected static Dimension dimension(Element e, Namespace xdr) {
        Element fromEle = e.element(QName.get("from", xdr));
        int[] f = dimEle(fromEle, xdr);
        Element toEle = e.element(QName.get("to", xdr));
        int[] t = dimEle(toEle, xdr);

        return new Dimension(f[0] + 1, (short) (f[1] + 1), t[0] + 1, (short) (t[1] + 1));
    }

    protected static int[] dimEle(Element e, Namespace xdr) {
        int c = 0, r = 0;
        if (e != null) {
            String col = e.element(QName.get("col", xdr)).getText(), row = e.element(QName.get("row", xdr)).getText();
            c = Integer.parseInt(col);
            r = Integer.parseInt(row);
        }
        return new int[] { r, c };
    }

    public List<Picture> getPictures() {
        return pictures;
    }

    public void setPictures(List<Picture> pictures) {
        this.pictures = pictures;
    }

    public boolean isParsed() {
        return parsed;
    }

    public void setParsed(boolean parsed) {
        this.parsed = parsed;
    }
}

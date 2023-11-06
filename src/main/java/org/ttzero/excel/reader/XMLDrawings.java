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
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.ttzero.excel.entity.Relationship;
import org.ttzero.excel.manager.Const;
import org.ttzero.excel.manager.RelManager;
import org.ttzero.excel.util.FileUtil;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;

import static org.ttzero.excel.reader.ExcelReader.getEntry;
import static org.ttzero.excel.reader.ExcelReader.toZipPath;

/**
 * 读取xlsx格式Excel图片，解析{@code drawing.xml}和{@code cellimages.xml}，
 * 后者是WPS自定义的嵌入图片，内嵌图片是整个工作薄全局共享的所有无法不包含单元格信息，
 * 为了和Excel图片图片统一接口需要先解析工作表然后再和内嵌图片的ID进行映射
 *
 * @author guanquan.wang at 2021-04-24 16:18
 */
public class XMLDrawings implements Drawings {
    /**
     * LOGGER
     */
    protected final Logger LOGGER = LoggerFactory.getLogger(getClass());
    /**
     * ExcelReader
     */
    protected final ExcelReader excelReader;
    /**
     * 临时保存所有工作表包含的图片
     */
    protected List<Drawings.Picture> pictures;
    /**
     * 是否已解析，保证数据只被解析一次
     */
    protected boolean parsed;

    public XMLDrawings(ExcelReader reader) {
        this.excelReader = reader;
    }

    /**
     * 列出所有工作表包含的图片
     *
     * @return 如果存在图片时返回 {@link Picture}数组, 不存在图片返回{@code null}.
     */
    @Override
    public List<Drawings.Picture> listPictures() {
        return parsed ? pictures : parse();
    }

    /**
     * 解析图片
     *
     * @return 列出所有图片 {@link Picture}
     */
    protected List<Drawings.Picture> parse() {
        parsed = true;
        // Empty excel, maybe throw exception here
        if (excelReader.sheets == null) return null;

        ZipFile zipFile = excelReader.zipFile;
        if (zipFile == null) return null;

        // 兼容读取WPS内嵌图片cellimages.xml
        ZipEntry cellImagesEntry = getEntry(zipFile, "xl/cellimages.xml");
        // 内嵌图片临时缓存 ID: 临时路径
        Map<String, Path> cellImagesMapper = cellImagesEntry != null ? listCellImages(zipFile, cellImagesEntry) : null;
        boolean hasCellImages = cellImagesMapper != null && !cellImagesMapper.isEmpty();

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
                    } catch (IOException ex) {
                        LOGGER.error("Copy image into {} failed", target, ex);
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

            // WPS内嵌图片兼容处理
            if (hasCellImages) {
                try {
                    pictures.addAll(quickFindCellImages(sheet, cellImagesMapper));
                } catch (IOException e) {
                    LOGGER.error("Parse build-in cell-images failed", e);
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

    /**
     * 拉取WPS单元格内嵌图片
     *
     * @return ID:图片本地路径
     */
    public Map<String, Path> listCellImages(ZipFile zipFile, ZipEntry entry) {
        SAXReader reader = SAXReader.createDefault();

        ZipEntry refEntry = getEntry(zipFile, "xl/_rels/cellimages.xml.rels");
        if (refEntry == null) return Collections.emptyMap();
        Document document;
        try {
            document = reader.read(zipFile.getInputStream(refEntry));
        } catch (DocumentException | IOException e) {
            LOGGER.warn("Read [xl/_rels/cellimages.xml.rels] failed.", e);
            return null;
        }
        List<Element> list = document.getRootElement().elements();
        Relationship[] rels = new Relationship[list.size()];
        int i = 0;
        for (Element e : list) {
            rels[i++] = new Relationship(e.attributeValue("Id"), e.attributeValue("Target"), e.attributeValue("Type"));
        }
        RelManager relManager = RelManager.of(rels);

        Element cellImages;
        try {
            cellImages = reader.read(zipFile.getInputStream(entry)).getRootElement();
        } catch (IOException | DocumentException e) {
            LOGGER.warn("Read [xl/cellimages.xml] failed.", e);
            return null;
        }
        List<Element> images = cellImages.elements();
        Namespace xdr = cellImages.getNamespaceForPrefix("xdr"), a = cellImages.getNamespaceForPrefix("a");
        // 图片临时存放的位置
        if (excelReader.tempDir == null) {
            try {
                excelReader.tempDir = FileUtil.mktmp("eec-");
            } catch (IOException e) {
                throw new ExcelReadException("创建临时文件夹失败.", e);
            }
        }
        Map<String, Path> cellImageMapper = new HashMap<>(images.size());
        for (Element e : images) {
            Element pic = e.element(QName.get("pic", xdr));
            // Not a picture
            if (pic == null) continue;

            Element nvPicPr = pic.element(QName.get("nvPicPr", xdr));
            if (nvPicPr == null) continue;
            Element cNvPr = nvPicPr.element(QName.get("cNvPr", xdr));
            if (cNvPr == null) continue;
            String name = cNvPr.attributeValue("name");

            Element blipFill = pic.element(QName.get("blipFill", xdr));
            if (blipFill == null) continue;

            Element blip = blipFill.element(QName.get("blip", a));
            if (blip == null) continue;

            Namespace r = blip.getNamespaceForPrefix("r");
            String embed = blip.attributeValue(QName.get("embed", r));
            Relationship rel = relManager.getById(embed);
            if (r != null && Const.Relationship.IMAGE.equals(rel.getType())) {
                Path localPath = null;
                // 复制图片到临时文件夹
                entry = getEntry(zipFile, "xl/" + rel.getTarget());
                if (entry != null) {
                    try {
                        Path targetPath = excelReader.tempDir.resolve(rel.getTarget());
                        if (!Files.exists(targetPath.getParent())) {
                            Files.createDirectories(targetPath);
                        }
                        Files.copy(zipFile.getInputStream(entry), targetPath, StandardCopyOption.REPLACE_EXISTING);
                        localPath = targetPath;
                    } catch (IOException ioException) {
                        ioException.printStackTrace();
                    }
                    cellImageMapper.put(name, localPath);
                }

            }
        }
        return cellImageMapper;
    }

    /**
     * 快整查询内嵌图片在工作表中的位置
     *
     * @param sheet 工作表
     * @param cellImageMapper 图片ID映射关系
     * @return 图片列表
     * @throws IOException if I/O error occur
     */
    protected List<Picture> quickFindCellImages(Sheet sheet, Map<String, Path> cellImageMapper) throws IOException {
        List<Picture> pictures = new ArrayList<>();
        String formula;
        // 转为CalcSheet工作表解析公式，解析类似<f>_xlfn.DISPIMG("图片ID",1)</f>，取出图片ID与Mapper进行匹配
        for (Iterator<Row> iter = sheet.asCalcSheet().load().iterator(); iter.hasNext(); ) {
            Row row = iter.next();
            for (int i = row.getFirstColumnIndex(), len = row.getLastColumnIndex(); i < len; i++) {
                if ((formula = row.getFormula(i)) != null && formula.startsWith("_xlfn.DISPIMG(\"")) {
                    formula = formula.substring(15, formula.lastIndexOf('"'));
                    Path path = cellImageMapper.get(formula);
                    if (path != null) {
                        Picture pic = new Picture();
                        pic.sheet = sheet;
                        pic.localPath = path;
                        pic.dimension = new Dimension(row.getRowNum(), (short) (i + 1), row.getRowNum(), (short) (i + 1));
                        pictures.add(pic);
                    }
                }
            }
        }
        return pictures;
    }
}

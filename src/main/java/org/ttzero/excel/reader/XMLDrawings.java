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
import org.ttzero.excel.entity.style.Styles;
import org.ttzero.excel.manager.Const;
import org.ttzero.excel.manager.RelManager;
import org.ttzero.excel.util.FileUtil;
import org.ttzero.excel.util.StringUtil;

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
 * 后者是WPS自定义的嵌入图片，内嵌图片是整个工作薄全局共享的所以无法不包含单元格信息，
 * 为了和Excel图片图片统一接口需要先解析工作表然后再和内嵌图片的ID进行映射，由于会对工作表
 * 进行两次读取所以对性能有一定影响，行数小于{@code 1}万影响不大可放心使用，当然你也可以直接
 * 调用本类的{@link #listCellImages(ZipFile, ZipEntry)}方法获取图片ID映射，然后在读取
 * 工作表时自己进行ID和单元格行列映射，这样做只会进行一次工作表读不会影响正常的读取性能。
 *
 * <p>参考文档:</p>
 * <p><a href="https://github.com/wangguanquan/eec/issues/363">解析POI内嵌图片</a></p>
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

        // FIXME 目前使用dom4j解析，如果批注较多时耗时和内存将增大增

        List<Picture> pictures = new ArrayList<>();
        for (Sheet sheet : excelReader.sheets) {
            XMLSheet xmlSheet = (XMLSheet) sheet;
            List<Relationship> list = xmlSheet.getRelManager().getAllByTypes(Const.Relationship.DRAWINGS, Const.Relationship.IMAGE);
            for (Relationship e : list) {
                String target = e.getTarget(), type = e.getType();
                ZipEntry entry = getEntry(zipFile, "xl/" + toZipPath(target));
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

            // 兼容读取WPS内嵌图片cellimages.xml
            ZipEntry cellImagesEntry = getEntry(zipFile, "xl/cellimages.xml");
            // 内嵌图片临时缓存 ID: 临时路径
            Map<String, Path> cellImagesMapper = cellImagesEntry != null ? listCellImages(zipFile, cellImagesEntry) : null;
            // WPS内嵌图片兼容处理
            if (cellImagesMapper != null && !cellImagesMapper.isEmpty()) {
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
        if (entry1 == null) return null; //throw new ExcelReadException("The file format is incorrect or corrupted. [" + key + "]");
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
        // Ignore hidden picture
        boolean ignoreIfHidden = ignoreHiddenPicture();
        Map<String, Path> localPathMap = new HashMap<>(Math.min(1 << 8, elements.size()));
        for (Element e : root.elements()) {
            Element pic = e.element(QName.get("pic", xdr));
            // Not a picture
            if (pic == null) continue;

            Element blipFill = pic.element(QName.get("blipFill", xdr));
            if (blipFill == null) continue;

            Element blip = blipFill.element(QName.get("blip", a));
            if (blip == null) continue;

            if (ignoreIfHidden) {
                Element nvPicPr = pic.element(QName.get("nvPicPr", xdr)), cNvPr;
                // Ignore if hidden
                if (nvPicPr != null && (cNvPr = nvPicPr.element(QName.get("cNvPr", xdr))) != null && "1".equals(cNvPr.attributeValue("hidden"))) {
                    continue;
                }
            }

            Namespace r = blip.getNamespaceForPrefix("r");
            String embed = blip.attributeValue(QName.get("embed", r));
            Relationship rel = relManager.getById(embed);
            if (rel == null || !Const.Relationship.IMAGE.equals(rel.getType())) continue;

            Picture picture = new Picture();
            pictures.add(picture);
            String target = toZipPath(rel.getTarget());
            Path targetPath = localPathMap.get(target);
            if (targetPath == null && (entry = getEntry(zipFile, "xl/" + target)) != null) {
                // Copy image to tmp path
                try {
                    targetPath = imagesPath.resolve(rel.getTarget());
                    Files.copy(zipFile.getInputStream(entry), targetPath, StandardCopyOption.REPLACE_EXISTING);
                    localPathMap.put(target, targetPath);
                } catch (IOException ex) { }
            }
            picture.localPath = targetPath;

            int[][] ft = parseDimension(e, xdr);
            boolean oneCellAnchor = "oneCellAnchor".equals(e.getName());
            if (oneCellAnchor) {
                picture.dimension = new Dimension(ft[0][2] + 1, (short) (ft[0][0] + 1));
            } else {
                picture.dimension = new Dimension(ft[0][2] + 1, (short) (ft[0][0] + 1), ft[1][2] + 1, (short) (ft[1][0] + 1));
            }
            picture.padding = new short[] { (short) ft[0][3], (short) ft[1][1], (short) ft[1][3], (short) ft[0][1] };
            String editAs = e.attributeValue("editAs");
            int property = -1;
            if (StringUtil.isNotEmpty(editAs)) {
                switch (editAs) {
                    case "twoCell" : property = 0; break;
                    case "oneCell" : property = 1; break;
                    case "absolute": property = 2; break;
                    default:
                }
            }
            picture.property = property | (oneCellAnchor ? 1 : 0) << 3;
            Element spPr = pic.element(QName.get("spPr", xdr));
            if (spPr != null) {
                Element xfrm = spPr.element(QName.get("xfrm", a));
                String rot;
                if (xfrm != null && StringUtil.isNotBlank(rot = xfrm.attributeValue("rot"))) {
                    try {
                        picture.revolve = Integer.parseInt(rot) / 60000;
                    } catch (Exception ex) {
                        // Ignore
                    }
                }

                // TODO Attach picture effects
            }

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
        return !pictures.isEmpty() ? pictures : null;
    }

    protected static int[][] parseDimension(Element e, Namespace xdr) {
        Element fromEle = e.element(QName.get("from", xdr));
        int[] f = dimEle(fromEle, xdr), t;
        if ("oneCellAnchor".equals(e.getName())) {
            Element ext = e.element(QName.get("ext", xdr));
            String cx = Styles.getAttr(ext, "cx"), cy = Styles.getAttr(ext, "cy");
            int width = StringUtil.isNotBlank(cx) ? Integer.parseInt(cx) : 0, height = StringUtil.isNotBlank(cy) ? Integer.parseInt(cy) : 0;
            t = new int[] { 0, width, 0, height};
        } else {
            Element toEle = e.element(QName.get("to", xdr));
            t = dimEle(toEle, xdr);
        }

        return new int[][] { f, t };
    }

    protected static int[] dimEle(Element e, Namespace xdr) {
        int c = 0, r = 0, co = 0, ro = 0;
        if (e != null) {
            String col = e.element(QName.get("col", xdr)).getText()
                , colOff = e.element(QName.get("colOff", xdr)).getText()
                , row = e.element(QName.get("row", xdr)).getText()
                , rowOff = e.element(QName.get("rowOff", xdr)).getText();
            c = Integer.parseInt(col);
            r = Integer.parseInt(row);
            co = (int) (Integer.parseInt(colOff) / 12700.0D + 0.5);
            ro = (int) (Integer.parseInt(rowOff) / 12700.0D + 0.5);
        }
        return new int[] { c, co, r, ro };
    }

    /**
     * 拉取WPS单元格内嵌图片
     *
     * @param zipFile xlsx源
     * @param entry   cellImages
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
                throw new ExcelReadException("Create temp directory failed.", e);
            }
        }
        Map<String, Path> cellImageMapper = new HashMap<>(Math.min(1 << 8, images.size()));
        // Ignore hidden picture
        boolean ignoreIfHidden = ignoreHiddenPicture();
        for (Element e : images) {
            Element pic = e.element(QName.get("pic", xdr));
            // Not a picture
            if (pic == null) continue;

            Element nvPicPr = pic.element(QName.get("nvPicPr", xdr));
            if (nvPicPr == null) continue;
            Element cNvPr = nvPicPr.element(QName.get("cNvPr", xdr));
            // Ignore if hidden
            if (cNvPr == null || ignoreIfHidden && "1".equals(cNvPr.attributeValue("hidden"))) continue;
            String name = cNvPr.attributeValue("name");

            Element blipFill = pic.element(QName.get("blipFill", xdr));
            if (blipFill == null) continue;

            Element blip = blipFill.element(QName.get("blip", a));
            if (blip == null) continue;

            Namespace r = blip.getNamespaceForPrefix("r");
            String embed = blip.attributeValue(QName.get("embed", r));
            Relationship rel = relManager.getById(embed);
            if (r == null || !Const.Relationship.IMAGE.equals(rel.getType())) continue;
            Path localPath = cellImageMapper.get(name), targetPath = excelReader.tempDir.resolve(rel.getTarget());
            if (localPath != null && localPath.equals(targetPath)) continue;
            // 复制图片到临时文件夹
            entry = getEntry(zipFile, "xl/" + rel.getTarget());
            if (entry != null) {
                try {
                    if (!Files.exists(targetPath.getParent())) {
                        Files.createDirectories(targetPath);
                    }
                    Files.copy(zipFile.getInputStream(entry), targetPath, StandardCopyOption.REPLACE_EXISTING);
                    localPath = targetPath;
                } catch (IOException ex) {
                    LOGGER.warn("Copy picture error.", ex);
                }
            }
            cellImageMapper.put(name, localPath);
        }
        return cellImageMapper;
    }

    /**
     * 快整查询内嵌图片在工作表中的位置
     *
     * @param sheet           工作表
     * @param cellImageMapper 图片ID映射关系
     * @return 图片列表
     * @throws IOException if I/O error occur
     */
    protected List<Picture> quickFindCellImages(Sheet sheet, Map<String, Path> cellImageMapper) throws IOException {
        List<Picture> pictures = new ArrayList<>();
        String formula;
        // 转为CalcSheet工作表解析公式，解析类似<f>_xlfn.DISPIMG("图片ID",1)</f>，取出图片ID与Mapper进行匹配
        for (Iterator<Row> iter = sheet.asFullSheet().load().iterator(); iter.hasNext(); ) {
            Row row = iter.next();
            for (int i = row.getFirstColumnIndex(), len = row.getLastColumnIndex(); i < len; i++) {
                if ((formula = row.getFormula(i)) != null && formula.startsWith("_xlfn.DISPIMG(\"")) {
                    formula = formula.substring(15, formula.lastIndexOf('"'));
                    Path path = cellImageMapper.get(formula);
                    if (path != null) {
                        Picture pic = new Picture();
                        pic.sheet = sheet;
                        pic.localPath = path;
                        pic.dimension = new Dimension(row.getRowNum(), (short) (i + 1), row.getRowNum() + 1, (short) (i + 2));
                        pic.padding = new short[] {1, -1, -1, 1};
                        pictures.add(pic);
                    }
                }
            }
        }
        return pictures;
    }

    /**
     * 是否忽略隐藏图片
     *
     * @return {@code true}忽略
     */
    protected boolean ignoreHiddenPicture() {
        return true;
    }
}

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

package org.ttzero.excel.entity.e7;

import org.dom4j.Document;
import org.dom4j.DocumentFactory;
import org.dom4j.Element;
import org.dom4j.Namespace;
import org.dom4j.QName;
import org.ttzero.excel.annotation.TopNS;
import org.ttzero.excel.entity.Comments;
import org.ttzero.excel.entity.ExcelWriteException;
import org.ttzero.excel.entity.IWorkbookWriter;
import org.ttzero.excel.entity.IWorksheetWriter;
import org.ttzero.excel.entity.Relationship;
import org.ttzero.excel.entity.SharedStrings;
import org.ttzero.excel.entity.Sheet;
import org.ttzero.excel.entity.WaterMark;
import org.ttzero.excel.entity.Workbook;
import org.ttzero.excel.entity.style.Fill;
import org.ttzero.excel.entity.style.PatternType;
import org.ttzero.excel.manager.Const;
import org.ttzero.excel.manager.RelManager;
import org.ttzero.excel.manager.docProps.App;
import org.ttzero.excel.manager.docProps.Core;
import org.ttzero.excel.util.FileUtil;
import org.ttzero.excel.util.StringUtil;
import org.ttzero.excel.util.ZipUtil;

import java.awt.Color;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Properties;

import static org.ttzero.excel.util.FileUtil.exists;

/**
 * @author guanquan.wang at 2019-04-22 15:47
 */
@TopNS(prefix = {"", "r"}, value = "workbook"
    , uri = {Const.SCHEMA_MAIN, Const.Relationship.RELATIONSHIP})
public class XMLWorkbookWriter implements IWorkbookWriter {

    private Workbook workbook;
    private final RelManager relManager;

    public XMLWorkbookWriter() {
        relManager = new RelManager();
    }

    public XMLWorkbookWriter(Workbook workbook) {
        this.workbook = workbook;
        relManager = new RelManager();
    }

    public Workbook getWorkbook() {
        return workbook;
    }

    public RelManager getRelManager() {
        return relManager;
    }

    /**
     * Setting workbook
     *
     * @param workbook the global workbook context
     */
    @Override
    public void setWorkbook(Workbook workbook) {
        this.workbook = workbook;
    }

    /**
     * The Workbook suffix
     *
     * @return xlsx if excel07, xls if excel03
     */
    @Override
    public String getSuffix() {
        return Const.Suffix.EXCEL_07;
    }

    /**
     * Write the workbook file to ${path}
     *
     * @param path the storage path
     */
    @Override
    public void writeTo(Path path) throws IOException {
        Path zip = workbook.getTemplate() == null ? createTemp() : template();
        reMarkPath(zip, path);
        FileUtil.rm(zip);
    }

    @Override
    public void writeTo(OutputStream os) throws IOException {
        Path zip = workbook.getTemplate() == null ? createTemp() : template();
        Files.copy(zip, os);
        FileUtil.rm(zip);
    }

    @Override
    public void writeTo(File file) throws IOException {
        Path zip = workbook.getTemplate() == null ? createTemp() : template();
        FileUtil.cp(zip, file);
        FileUtil.rm(zip);
    }


    // --- PRIVATE FUNCTIONS


    private void addRel(Relationship rel) {
        relManager.add(rel);
    }

    private void writeXML(Path root) throws IOException {

        // Content type
        ContentType contentType = new ContentType();
        contentType.add(new ContentType.Default(Const.ContentType.RELATIONSHIP, "rels"));
        contentType.add(new ContentType.Default(Const.ContentType.XML, "xml"));
        contentType.add(new ContentType.Override(Const.ContentType.SHAREDSTRING, "/xl/sharedStrings.xml"));
        contentType.add(new ContentType.Override(Const.ContentType.WORKBOOK, "/xl/workbook.xml"));
        contentType.addRel(new Relationship("xl/workbook.xml", Const.Relationship.OFFICE_DOCUMENT));

        // Write app
        writeApp(root, contentType);

        // Write core
        writeCore(root, contentType);

        Path themeP = root.resolve("theme");
        if (!exists(themeP)) {
            Files.createDirectory(themeP);
        }
        try {
            InputStream theme = getClass().getClassLoader().getResourceAsStream("template/theme1.xml");
            if (theme != null) {
                Files.copy(theme, themeP.resolve("theme1.xml"));
            }
        } catch (IOException e) {
            // Nothing
        }
        addRel(new Relationship("theme/theme1.xml", Const.Relationship.THEME));
        contentType.add(new ContentType.Override(Const.ContentType.THEME, "/xl/theme/theme1.xml"));

        // style
        addRel(new Relationship("styles.xml", Const.Relationship.STYLE));
        contentType.add(new ContentType.Override(Const.ContentType.STYLE, "/xl/styles.xml"));

        addRel(new Relationship("sharedStrings.xml", Const.Relationship.SHARED_STRING));

        WaterMark waterMark;
        if ((waterMark = workbook.getWaterMark()) != null) {
            contentType.add(new ContentType.Default(waterMark.getContentType(), waterMark.getSuffix().substring(1)));
        }

        int size = workbook.getSize();
        for (int i = 0; i < size; i++) {
            WaterMark wm = workbook.getSheetAt(i).getWaterMark();
            if (wm != null) {
                contentType.add(new ContentType.Default(wm.getContentType(), wm.getSuffix().substring(1)));
            }
        }

        for (int i = 0; i < size; i++) {
            Sheet sheet;
            contentType.add(new ContentType.Override(Const.ContentType.SHEET
                , "/xl/worksheets/sheet" + (sheet = workbook.getSheetAt(i)).getId() + Const.Suffix.XML));
            Comments comments = sheet.getComments();
            if (comments != null) {
                comments.writeTo(root);
                contentType.add(new ContentType.Override(Const.ContentType.COMMENTS
                    , "/xl/comments" + sheet.getId() + Const.Suffix.XML));
                contentType.add(new ContentType.Default(Const.ContentType.VMLDRAWING, "vml"));
            }
        } // END


        // write content type
        contentType.writeTo(root.getParent());

        TopNS topNS = getClass().getAnnotation(TopNS.class);
        String name;
        if (topNS != null) {
            name = topNS.value();
        } else name = "workbook";
        // Relationship
        relManager.write(root, name + Const.Suffix.XML);

        // workbook.xml
        writeSelf(root);

        // styles
        workbook.getStyles().writeTo(root.resolve("styles.xml"));

        // share string
        try (SharedStrings sst = workbook.getSst()) {
            sst.writeTo(root);
        }
    }

    private void writeApp(Path root, ContentType contentType) throws IOException {

        // docProps
        App app = new App();
        // Write company name if set
        if (StringUtil.isNotEmpty(workbook.getCompany())) {
            app.setCompany(workbook.getCompany());
        }

        // Read app and version from pom
        Properties pom = IWorkbookWriter.pom();

        app.setApplication(pom.getProperty("groupId") + "." + pom.getProperty("artifactId"));
        app.setAppVersion(pom.getProperty("version"));

        int size = workbook.getSize();

        List<String> titleParts = new ArrayList<>(size);
        for (int i = 0; i < size; i++) {
            Sheet sheet = workbook.getSheetAt(i);
            titleParts.add(sheet.getName());
            addRel(new Relationship("worksheets/sheet" + sheet.getId() + Const.Suffix.XML, Const.Relationship.SHEET));
        }
        app.setTitlePards(titleParts);

        app.writeTo(root.getParent().resolve("docProps/app.xml"));
        contentType.add(new ContentType.Override(Const.ContentType.APP, "/docProps/app.xml"));
        contentType.addRel(new Relationship("docProps/app.xml", Const.Relationship.APP));
    }

    private void writeCore(Path root, ContentType contentType) throws IOException {
        Core core = workbook.getCore() != null ? workbook.getCore() : new Core();
        if (StringUtil.isEmpty(core.getCreator())) {
            if (workbook.getCreator() != null) {
                core.setCreator(workbook.getCreator());
            } else {
                core.setCreator(System.getProperty("user.name"));
            }
        }
        if (StringUtil.isEmpty(core.getTitle())) core.setTitle(workbook.getName());
        if (core.getCreated() == null) core.setCreated(new Date());
        if (core.getModified() == null) core.setModified(new Date());

        core.writeTo(root.getParent().resolve("docProps/core.xml"));
        contentType.add(new ContentType.Override(Const.ContentType.CORE, "/docProps/core.xml"));
        contentType.addRel(new Relationship("docProps/core.xml", Const.Relationship.CORE));
    }

    private void madeMark(Path parent) throws IOException {
        Relationship supRel = null;
        int n = 1;
        WaterMark waterMark = workbook.getWaterMark();
        if (waterMark != null) {
            Path media = parent.resolve("media");
            if (!exists(media)) {
                Files.createDirectory(media);
            }
            Path image = media.resolve("image" + n++ + waterMark.getSuffix());

            Files.copy(waterMark.get(), image);
            supRel = new Relationship("../media/" + image.getFileName(), Const.Relationship.IMAGE);
        }
        WaterMark wm;
        for (int i = 0; i < workbook.getSize(); i++) {
            Sheet sheet = workbook.getSheetAt(i);
            if ((wm = sheet.getWaterMark()) != null) {
                Path media = parent.resolve("media");
                if (!exists(media)) {
                    Files.createDirectory(media);
                }
                Path image = media.resolve("image" + n++ + wm.getSuffix());
                Files.copy(wm.get(), image);
                sheet.addRel(new Relationship("../media/" + image.getFileName(), Const.Relationship.IMAGE));
            } else if (waterMark != null) {
                sheet.setWaterMark(waterMark);
                sheet.addRel(supRel);
            }
        }
    }

    private void writeSelf(Path root) throws IOException {
        DocumentFactory factory = DocumentFactory.getInstance();
        //use the factory to create a root element
        Element rootElement = null;
        //use the factory to create a new document with the previously created root element
        boolean hasTopNs;
        String[] prefixs = null, uris = null;
        String rootName = null;
        TopNS topNs = getClass().getAnnotation(TopNS.class);
        if (hasTopNs = getClass().isAnnotationPresent(TopNS.class)) {
            prefixs = topNs.prefix();
            uris = topNs.uri();
            rootName = topNs.value();
            for (int i = 0; i < prefixs.length; i++) {
                if (prefixs[i].length() == 0) {
                    rootElement = factory.createElement(rootName, uris[i]);
                    break;
                }
            }
        }
        if (rootElement == null) {
            if (hasTopNs) {
                rootElement = factory.createElement(rootName);
            } else {
                workbook.what("9004", "workbook.xml");
                return;
            }
        }

        if (hasTopNs) {
            for (int i = 0; i < prefixs.length; i++) {
                rootElement.add(Namespace.get(prefixs[i], uris[i]));
            }
        }

        // book view
        rootElement.addElement("bookViews").addElement("workbookView").addAttribute("activeTab", "0");

        // sheets
        Element sheetEle = rootElement.addElement("sheets");
        for (int i = 0; i < workbook.getSize(); i++) {
            Sheet sheetInfo = workbook.getSheetAt(i);
            Element st = sheetEle.addElement("sheet")
                .addAttribute("sheetId", String.valueOf(i + 1))
                .addAttribute("name", sheetInfo.getName());
            if (sheetInfo.isHidden()) {
                st.addAttribute("state", "hidden");
            }
            Relationship rs = relManager.getByTarget("worksheets/sheet" + (i + 1) + Const.Suffix.XML);
            if (rs != null) {
                st.addAttribute(QName.get("id", Namespace.get("r", uris[StringUtil.indexOf(prefixs, "r")])), rs.getId());
            }
        }

        Document doc = factory.createDocument(rootElement);
        FileUtil.writeToDiskNoFormat(doc, root.resolve(rootName + Const.Suffix.XML)); // write to desk
    }

    //////////////////////////////////////////////////////
    protected Path createTemp() throws IOException, ExcelWriteException {
        Sheet[] sheets = workbook.getSheets();
        for (int i = 0; i < sheets.length; i++) {
            Sheet sheet = sheets[i];
            if (sheet.getSheetWriter() == null) {
                IWorksheetWriter worksheetWriter = getWorksheetWriter(sheet);
                sheet.setSheetWriter(worksheetWriter);
            }
            if (sheet.getAutoSize() == 0) {
                if (workbook.isAutoSize()) {
                    sheet.autoSize();
                } else {
                    sheet.fixSize();
                }
            }
            if (sheet.getAutoOdd() == -1) {
                sheet.setAutoOdd(workbook.getAutoOdd());
            }
            // Interlace changes color as default
            if (sheet.getAutoOdd() == 0) {
                sheet.setOddFill(workbook.getOddFill() == null
                    ? new Fill(PatternType.solid, new Color(239, 245, 235)) : workbook.getOddFill());
            }
            sheet.setId(i + 1);
            // default worksheet name
            if (StringUtil.isEmpty(sheet.getName())) {
                sheet.setName("Sheet" + (i + 1));
            }
            // Set cell value and style processor
            sheet.setCellValueAndStyle(new XMLCellValueAndStyle(sheet.getAutoOdd(), sheet.getOddFill()));

            // Force export all fields
            if (workbook.getForceExport() > sheet.getForceExport()) {
                sheet.forceExport();
            }
        }
        workbook.what("0001");

        Path root = null;
        try {
            root = FileUtil.mktmp(Const.EEC_PREFIX);
            workbook.what("0002", root.toString());

            Path xl = Files.createDirectory(root.resolve("xl"));
            // Create  watermark first, it need to use when writing each sheet
            madeMark(xl);

            // Write worksheet data one by one
            for (int i = 0; i < workbook.getSize(); i++) {
                Sheet e = workbook.getSheetAt(i);
                e.writeTo(xl);
                if (e.getWaterMark() != null)
                    e.getWaterMark().delete(); // Delete template image
                e.close();
            }

            // Write SharedString, Styles and workbook.xml
            writeXML(xl);
            if (workbook.getWaterMark() != null)
                workbook.getWaterMark().delete() ; // Delete template image
            workbook.what("0003");

            // Zip compress
            Path zipFile = ZipUtil.zipExcludeRoot(root, root);
            workbook.what("0004", zipFile.toString());

            // Delete source files
            FileUtil.rm_rf(root.toFile(), true);
            workbook.what("0005");
            return zipFile;
        } catch (IOException | ExcelWriteException e) {
            // remove temp path
            if (root != null) FileUtil.rm_rf(root);
            workbook.getSst().close();
            throw e;
        }
    }

    protected void reMarkPath(Path zip, Path path) throws IOException {
        String name;
        if (StringUtil.isEmpty(name = workbook.getName())) {
            name = workbook.getI18N().getOrElse("non-name-file", "Non name");
        }

        Path resultPath = reMarkPath(zip, path, name);
        workbook.what("0006", resultPath.toString());
    }

    // --- TEMPLATE

    @Override
    public Path template() throws IOException {
        workbook.what("0007");
        // Store template stream as zip file
        Path temp = FileUtil.mktmp(Const.EEC_PREFIX);
        ZipUtil.unzip(workbook.getTemplate(), temp);
        workbook.what("0008");

        // Bind data
        EmbedTemplate bt = new EmbedTemplate(temp, workbook);
        if (bt.check()) { // Check files
            bt.bind(workbook.getBind());
        }
        workbook.what("0003");

        // Zip compress
        Path zipFile = ZipUtil.zipExcludeRoot(temp, temp);
        workbook.what("0004", zipFile.toString());

        // Delete source files
        FileUtil.rm_rf(temp.toFile(), true);
        workbook.what("0005");

        // Close shared string table
        workbook.getSst().close();

        return zipFile;
    }

    // --- Customize worksheet writer

    protected IWorksheetWriter getWorksheetWriter(Sheet sheet) {
        return new XMLWorksheetWriter(sheet);
    }
}

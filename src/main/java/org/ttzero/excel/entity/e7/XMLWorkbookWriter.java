/*
 * Copyright (c) 2017-2019, guanquan.wang@hotmail.com All Rights Reserved.
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
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.ttzero.excel.entity.ListSheet;
import org.ttzero.excel.manager.TopNS;
import org.ttzero.excel.entity.Comments;
import org.ttzero.excel.entity.ExcelWriteException;
import org.ttzero.excel.entity.ICellValueAndStyle;
import org.ttzero.excel.entity.IWorkbookWriter;
import org.ttzero.excel.entity.IWorksheetWriter;
import org.ttzero.excel.entity.Relationship;
import org.ttzero.excel.entity.SharedStrings;
import org.ttzero.excel.entity.Sheet;
import org.ttzero.excel.entity.Watermark;
import org.ttzero.excel.entity.Workbook;
import org.ttzero.excel.manager.Const;
import org.ttzero.excel.manager.RelManager;
import org.ttzero.excel.manager.docProps.App;
import org.ttzero.excel.manager.docProps.Core;
import org.ttzero.excel.manager.docProps.CustomProperties;
import org.ttzero.excel.util.FileUtil;
import org.ttzero.excel.util.StringUtil;
import org.ttzero.excel.util.ZipUtil;

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
    /**
     * LOGGER
     */
    protected final Logger LOGGER = LoggerFactory.getLogger(getClass());
    protected Workbook workbook;
    protected final RelManager relManager;

    public XMLWorkbookWriter() {
        relManager = new RelManager();
    }

    public XMLWorkbookWriter(Workbook workbook) {
        this.workbook = workbook;
        this.relManager = new RelManager();
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
        Path zip = null;
        try {
            zip = createTemp();
            moveToPath(zip, path);
        } finally {
            if (zip != null) FileUtil.rm(zip);
        }
    }

    @Override
    public void writeTo(OutputStream os) throws IOException {
        Path zip = null;
        try {
            zip = createTemp();
            Files.copy(zip, os);
        } finally {
            if (zip != null) FileUtil.rm(zip);
        }
    }


    // --- PRIVATE FUNCTIONS

    protected void writeGlobalAttribute(Path root) throws IOException {

        // Content type
        ContentType contentType = workbook.getContentType();
        contentType.add(new ContentType.Default(Const.ContentType.RELATIONSHIP, "rels"));
        contentType.add(new ContentType.Default(Const.ContentType.XML, "xml"));
        contentType.add(new ContentType.Override(Const.ContentType.WORKBOOK, "/xl/workbook.xml"));
        contentType.addRel(new Relationship("xl/workbook.xml", Const.Relationship.OFFICE_DOCUMENT));

        // Write app
        writeApp(root);

        // Write core
        writeCore(root);

        // Write custom properties
        writeCustomProperties(root);

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
        relManager.add(new Relationship("theme/theme1.xml", Const.Relationship.THEME));
        contentType.add(new ContentType.Override(Const.ContentType.THEME, "/xl/theme/theme1.xml"));

        // workbook.xml
        writeWorkbook(root);

        // styles
        workbook.getStyles().writeTo(root.resolve("styles.xml"));
        // style relationship
        relManager.add(new Relationship("styles.xml", Const.Relationship.STYLE));
        contentType.add(new ContentType.Override(Const.ContentType.STYLE, "/xl/styles.xml"));

        // share string
        try (SharedStrings sst = workbook.getSharedStrings()) {
            sst.writeTo(root);
        }
        relManager.add(new Relationship("sharedStrings.xml", Const.Relationship.SHARED_STRING));
        contentType.add(new ContentType.Override(Const.ContentType.SHAREDSTRING, "/xl/sharedStrings.xml"));

        // write content type
        contentType.writeTo(root.getParent());

        TopNS topNS = getClass().getAnnotation(TopNS.class);
        String name;
        if (topNS != null) {
            name = topNS.value();
        } else name = "workbook";
        // Relationship
        relManager.write(root, name + Const.Suffix.XML);
    }

    protected void writeApp(Path root) throws IOException {

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
            relManager.add(new Relationship("worksheets/sheet" + sheet.getId() + Const.Suffix.XML, Const.Relationship.SHEET));
        }
        app.setTitlePards(titleParts);

        app.writeTo(root.getParent().resolve("docProps/app.xml"));
        workbook.addContentType(new ContentType.Override(Const.ContentType.APP, "/docProps/app.xml"))
            .addContentTypeRel(new Relationship("docProps/app.xml", Const.Relationship.APP));
    }

    protected void writeCore(Path root) throws IOException {
        Core core = workbook.getCore() != null ? workbook.getCore() : new Core();
        if (StringUtil.isEmpty(core.getCreator()) && StringUtil.isNotEmpty(workbook.getCreator())) {
            core.setCreator(workbook.getCreator());
        }
        if (StringUtil.isEmpty(core.getTitle())) core.setTitle(workbook.getName());
        if (core.getCreated() == null) core.setCreated(new Date());
        if (core.getModified() == null) core.setModified(new Date());

        core.writeTo(root.getParent().resolve("docProps/core.xml"));
        workbook.addContentType(new ContentType.Override(Const.ContentType.CORE, "/docProps/core.xml"))
            .addContentTypeRel(new Relationship("docProps/core.xml", Const.Relationship.CORE));
    }

    /**
     * 写自定义属性
     *
     * @param root root路径
     * @throws IOException If I/O error occur.
     */
    protected void writeCustomProperties(Path root) throws IOException {
        CustomProperties custom = workbook.getCustomProperties();
        if (custom != null && custom.hasProperty()) {
            custom.writeTo(root.getParent().resolve("docProps/custom.xml"));
            workbook.addContentType(new ContentType.Override(Const.ContentType.CUSTOM, "/docProps/custom.xml"))
                .addContentTypeRel(new Relationship("docProps/custom.xml", Const.Relationship.CUSTOM));
        }
    }

    /**
     * @param root 根目录路径
     * @throws IOException 如果写入过程中发生I/O错误
     * @deprecated Rename to {@link #writeWorkbook(Path)}
     */
    @Deprecated
    protected void writeSelf(Path root) throws IOException {
        writeWorkbook(root);
    }

    /**
     * 将工作簿写入到指定路径
     *
     * @param root 根目录路径
     * @throws IOException 如果写入过程中发生I/O错误
     */
    protected void writeWorkbook(Path root) throws IOException {
        DocumentFactory factory = DocumentFactory.getInstance();
        //use the factory to create a root element
        Element rootElement = null;
        //use the factory to create a new document with the previously created root element
        String[] prefixs = null, uris = null;
        String rootName = null;
        TopNS topNs = getClass().getAnnotation(TopNS.class);
        boolean hasTopNs = (topNs != null);
        if (hasTopNs) {
            prefixs = topNs.prefix();
            uris = topNs.uri();
            rootName = topNs.value();
            for (int i = 0; i < prefixs.length; i++) {
                if (prefixs[i].isEmpty()) {
                    rootElement = factory.createElement(rootName, uris[i]);
                    break;
                }
            }
        }
        if (rootElement == null) {
            if (hasTopNs) {
                rootElement = factory.createElement(rootName);
            } else {
                LOGGER.error("Workbook missing necessary information.");
                return;
            }
        }

        if (prefixs.length > 0) {
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
                .addAttribute("sheetId", String.valueOf(sheetInfo.getId()))
                .addAttribute("name", sheetInfo.getName());
            if (sheetInfo.isHidden()) {
                st.addAttribute("state", "hidden");
            }
            Relationship rs = relManager.getByTarget("worksheets/sheet" + sheetInfo.getId() + Const.Suffix.XML);
            if (rs != null) {
                st.addAttribute(QName.get("id", Namespace.get("r", uris[StringUtil.indexOf(prefixs, "r")])), rs.getId());
            }
        }

        Document doc = factory.createDocument(rootElement);
        FileUtil.writeToDiskNoFormat(doc, root.resolve(rootName + Const.Suffix.XML)); // write to desk
    }

    protected void writeWorksheets(Path root) throws IOException {
        LOGGER.debug("Start to write Sheet.");
        ContentType contentType = workbook.getContentType();
        for (int i = 0; i < workbook.getSize(); i++) {
            Sheet sheet = workbook.getSheetAt(i);
            sheet.setId(i + 1);
            // default worksheet name
            if (StringUtil.isEmpty(sheet.getName())) {
                sheet.setName("Sheet" + sheet.getId());
            }
            if (sheet.getSheetWriter() == null) {
                sheet.setSheetWriter(getWorksheetWriter(sheet));
            }
            if (workbook.isAutoSize() && sheet.getAutoSize() == 0) {
                sheet.autoSize();
            }
            if (workbook.getZebraFill() != null && sheet.getZebraFillStyle() < 0) {
                sheet.setZebraLine(workbook.getZebraFill());
            }
            // Set cell value and style processor
            if (sheet.getCellValueAndStyle() == null) {
                int zebraFillStyle = sheet.getZebraFillStyle();
                ICellValueAndStyle cvas = zebraFillStyle > 0 ? new XMLZebraLineCellValueAndStyle(zebraFillStyle) : new XMLCellValueAndStyle();
                sheet.setCellValueAndStyle(cvas);
            }

            // Force export all fields
            if (workbook.getForceExport() > sheet.getForceExport() && ListSheet.class.isAssignableFrom(sheet.getClass())) {
                ((ListSheet<?>) sheet).forceExport();
            }

            // Merge Progress window
            if (workbook.getProgressConsumer() != null && sheet.getProgressConsumer() == null) {
                sheet.onProgress(workbook.getProgressConsumer());
            }

            try {
                // Write to desk
                sheet.writeTo(root);
            } finally {
                sheet.close();
            }

            // Add content-type
            contentType.add(new ContentType.Override(Const.ContentType.SHEET, "/xl/worksheets/sheet" + sheet.getId() + Const.Suffix.XML));

            // Add comments
            Comments comments = sheet.getComments();
            if (comments != null) {
                comments.writeTo(root);
                contentType.add(new ContentType.Override(Const.ContentType.COMMENTS, "/xl/comments" + sheet.getId() + Const.Suffix.XML));
                contentType.add(new ContentType.Default(Const.ContentType.VMLDRAWING, "vml"));
            }

            // Add watermark
            Watermark wm = sheet.getWatermark();
            if (wm != null && wm.canWrite()) {
                contentType.add(new ContentType.Default(wm.getContentType(), wm.getSuffix().substring(1)));
            }
        }
    }

    protected Path createTemp() throws IOException, ExcelWriteException {
        Path root = null;
        try {
            root = FileUtil.mktmp(Const.EEC_PREFIX);
            LOGGER.debug("Create temporary folder {}", root);

            Path xl = Files.createDirectory(root.resolve("xl"));

            // Write worksheet data one by one
            writeWorksheets(xl);

            // Write SharedString, Styles and workbook.xml
            writeGlobalAttribute(xl);
            LOGGER.debug("All sheets have completed writing, starting to compression ...");

            // Zip compress
            Path zipFile = ZipUtil.zipExcludeRoot(root, workbook.getCompressionLevel(), root);
            LOGGER.debug("Compression completed. {}", zipFile);

            return zipFile;
        } finally {
            // Remove temp path
            if (root != null) FileUtil.rm_rf(root);
        }
    }

    @Deprecated
    protected void reMarkPath(Path source, Path target) throws IOException {
        moveToPath(source, target);
    }

    protected void moveToPath(Path source, Path target) throws IOException {
        String name = StringUtil.isEmpty(workbook.getName()) ? "新建文件" : workbook.getName();
        Path resultPath = moveToPath(source, target, name);
        LOGGER.debug("Write completed. {}", resultPath);
    }

    @Override
    public void close() throws IOException {
        for (Sheet sheet : workbook.getSheets()) {
            if (sheet != null && sheet.getWatermark() != null)
                sheet.getWatermark().delete();
        }
        if (workbook.getWatermark() != null) workbook.getWatermark().delete() ; // Delete template image
        workbook.getSharedStrings().close();
    }

    // --- Customize worksheet writer

    public IWorksheetWriter getWorksheetWriter(Sheet sheet) {
        return new XMLWorksheetWriter(sheet);
    }
}

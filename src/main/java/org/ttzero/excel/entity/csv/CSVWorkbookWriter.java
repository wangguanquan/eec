/*
 * Copyright (c) 2019, guanquan.wang@yandex.com All Rights Reserved.
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

package org.ttzero.excel.entity.csv;

import org.ttzero.excel.entity.ExcelWriteException;
import org.ttzero.excel.entity.IWorkbookWriter;
import org.ttzero.excel.entity.IWorksheetWriter;
import org.ttzero.excel.entity.Sheet;
import org.ttzero.excel.entity.Workbook;
import org.ttzero.excel.manager.Const;
import org.ttzero.excel.util.FileUtil;
import org.ttzero.excel.util.StringUtil;
import org.ttzero.excel.util.ZipUtil;

import java.io.File;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;

/**
 * Write data as Comma-Separated Values format
 *
 * Create by guanquan.wang at 2019-08-21 21:46
 */
public class CSVWorkbookWriter implements IWorkbookWriter {
    private Workbook workbook;

    public CSVWorkbookWriter() { }

    public CSVWorkbookWriter(Workbook workbook) {
        this.workbook = workbook;
    }

    @Override
    public void setWorkbook(Workbook workbook) {
        this.workbook = workbook;
    }

    /**
     * The Comma-Separated Values format suffix
     *
     * @return const value '.csv'
     */
    @Override
    public String getSuffix() {
        return Const.Suffix.CSV;
    }

    /**
     * Write to OutputStream
     *
     * @param os the out put stream
     * @throws IOException         if io error occur
     */
    @Override
    public void writeTo(OutputStream os) throws IOException {
        Files.copy(createTemp(), os);
    }

    /**
     * Write to file
     *
     * @param file the storage file
     * @throws IOException         if io error occur
     */
    @Override
    public void writeTo(File file) throws IOException {
        FileUtil.cp(createTemp(), file);
    }

    /**
     * The Comma-Separated Values format do not support.
     *
     * @return the template path
     * @throws UnsupportedOperationException not support
     */
    @Override
    public Path template() {
        throw new UnsupportedOperationException();
    }

    @Override
    public void writeTo(Path root) throws IOException {

    }

    // Create csv file
    protected Path createTemp() throws IOException, ExcelWriteException {
        Sheet[] sheets = workbook.getSheets();
        for (int i = 0; i < sheets.length; i++) {
            Sheet sheet = sheets[i];
            IWorksheetWriter worksheetWriter = getWorksheetWriter(sheet);
            sheet.setSheetWriter(worksheetWriter);
            sheet.setId(i + 1);
            // default worksheet name
            if (StringUtil.isEmpty(sheet.getName())) {
                sheet.setName("Sheet" + (i + 1));
            }
        }
        workbook.what("0001");

        Path root = null;
        try {
            root = FileUtil.mktmp(Const.EEC_PREFIX);
            workbook.what("0002", root.toString());

            // Write worksheet data one by one
            for (int i = 0; i < workbook.getSize(); i++) {
                Sheet e = workbook.getSheetAt(i);
                e.writeTo(root);
                e.close();
            }

            // Zip compress if multi worksheet occur
            if (workbook.getSize() > 0) {
                Path zipFile = ZipUtil.zipExcludeRoot(root, root);
                workbook.what("0004", zipFile.toString());
                return zipFile;
            } else {
                return root.resolve(workbook.getSheetAt(0).getFileName());
            }
        } catch (IOException | ExcelWriteException e) {
            // remove temp path
            if (root != null) FileUtil.rm_rf(root);
            workbook.getSst().close();
            throw e;
        }
    }

    // --- Customize worksheet writer

    protected IWorksheetWriter getWorksheetWriter(Sheet sheet) {
        return new CSVWorksheetWriter(sheet);
    }
}

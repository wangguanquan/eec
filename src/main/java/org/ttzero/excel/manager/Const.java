/*
 * Copyright (c) 2017, guanquan.wang@yandex.com All Rights Reserved.
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

package org.ttzero.excel.manager;

/**
 * The Const class
 *
 * @author guanquan.wang on 2017/9/30.
 */
public interface Const {
    /**
     * Open xml schema
     */
    String SCHEMA_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    /**
     * Xml declaration
     */
    String XML_DECLARATION = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>";
    /**
     * Excel xml declatation
     */
    String EXCEL_XML_DECLARATION = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>";
    /**
     * "\n" in UNIX systems, "\r\n" in Windows systems.
     */
    String lineSeparator = System.lineSeparator();
    /**
     * Prefix of eec project
     */
    String EEC_PREFIX = "eec+";
    /**
     * Size of row-block
     */
    int ROW_BLOCK_SIZE = 32;

    /**
     * Relation
     */
    interface Relationship {
        String
            IMAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
            , APP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"
            , CORE = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"
            , CUSTOM = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties"
            , OFFICE_DOCUMENT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
            , SHARED_STRING = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"
            , STYLE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
            , SHEET = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
            , THEME = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
            , RELATIONSHIP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            , COMMENTS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
            , DRAWINGS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"
            , VMLDRAWING = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing"
            , CHARTS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"
            , HYPERLINK = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
            ;
    }

    /**
     * Content-type
     */
    interface ContentType {
        String
            PNG = "image/png"
            , JPG = "image/jpeg"
            , JPEG = "image/jpeg"
            , BMP = "image/bmp"
            , GIF = "image/gif"
            , XML = "application/xml"
            , THEME = "application/vnd.openxmlformats-officedocument.theme+xml"
            , STYLE = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"
            , RELATIONSHIP = "application/vnd.openxmlformats-package.relationships+xml"
            , WORKBOOK = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
            , SHEET = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
            , APP = "application/vnd.openxmlformats-officedocument.extended-properties+xml"
            , CORE = "application/vnd.openxmlformats-package.core-properties+xml"
            , CUSTOM = "application/vnd.openxmlformats-officedocument.custom-properties+xml"
            , SHAREDSTRING = "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"
            , PRINTSETTING = "application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings"
            , COMMENTS ="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"
            , DRAWINGS = "application/vnd.openxmlformats-officedocument.drawing+xml"
            , VMLDRAWING = "application/vnd.openxmlformats-officedocument.vmlDrawing"
            , CHARTS = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"
            ;
    }

    /**
     * Excel Limit
     */
    interface Limit {
        /**
         * Excel07's max rows on sheet
         */
        int MAX_ROWS_ON_SHEET = 1_048_576;
        /**
         * The max columns on sheet
         */
        int MAX_COLUMNS_ON_SHEET = 16_384;
        /**
         * The max characters per cell
         */
        int MAX_CHARACTERS_PER_CELL = 32_767;
        /**
         * The max line feeds per cell
         */
        int MAX_LINE_FEEDS_PER_CELL = 253;
        /**
         * Column width
         */
        int COLUMN_WIDTH = 255;
        /**
         * The header sub-columns length
         */
        int HEADER_SUB_COLUMNS = 1024;
    }

    /**
     * The file suffix
     */
    interface Suffix {
        /**
         * Excel 07
         */
        String EXCEL_07 = ".xlsx";
        /**
         * Excel 03
         */
        String EXCEL_03 = ".xls";
        /**
         * Xml
         */
        String XML = ".xml";
        /**
         * Relation
         */
        String RELATION = ".rels";
        /**
         * Png
         */
        String PNG = ".png";
        /**
         * CSV
         */
        String CSV = ".csv";
        /**
         * Zip
         */
        String ZIP = ".zip";
        /**
         * VML
         */
        String VML = ".vml";
    }

    /**
     * The extend properties key
     */
    interface ExtendPropertyKey {
        /**
         * Freeze panes key(0-position)
         */
        String FREEZE = "freeze";
        /**
         * The style design key(1-position)
         */
        String STYLE_DESIGN = "style_design";
        /**
         * Merge cells key(2-position)
         */
        String MERGE_CELLS  = "merge_cells";
        /**
         * AutoFilter(3-position)
         */
        String AUTO_FILTER = "auto_filter";
        /**
         * DataValidation(4-position)
         */
        String DATA_VALIDATION = "data_validation";
        /**
         * Zoom Scale(5-position)
         */
        String ZOOM_SCALE = "zoom_scale";
    }
}

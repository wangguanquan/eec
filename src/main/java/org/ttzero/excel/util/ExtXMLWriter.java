/*
 * Copyright (c) 2017, guanquan.wang@hotmail.com All Rights Reserved.
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

package org.ttzero.excel.util;

import org.dom4j.io.OutputFormat;
import org.dom4j.io.XMLWriter;

import java.io.IOException;
import java.io.OutputStream;
import java.io.UnsupportedEncodingException;

/**
 * @author guanquan.wang on 2017/10/11.
 */
public class ExtXMLWriter extends XMLWriter {

    public ExtXMLWriter(OutputStream out) throws UnsupportedEncodingException {
        super(out);
    }

    public ExtXMLWriter(OutputStream out, OutputFormat format)
        throws UnsupportedEncodingException {
        super(out, format);
    }

    /**
     * <p>
     * This will write the declaration to the given Writer. Assumes XML version
     * 1.0 since we don't directly know.
     * </p>
     *
     * @throws IOException DOCUMENT ME!
     */
    @Override
    protected void writeDeclaration() throws IOException {
        OutputFormat format = getOutputFormat();

        // Only print of declaration is not suppressed
        if (!format.isSuppressDeclaration()) {
            writer.write("<?xml version=\"1.0");
            if (!format.isOmitEncoding()) {
                writer.write("\" encoding=\"" + (StringUtil.isNotEmpty(format.getEncoding()) ? format.getEncoding() : "UTF-8"));
            }
            writer.write("\" standalone=\"yes\"?>");

            if (format.isNewLineAfterDeclaration()) {
                println();
            }
        }
    }
}

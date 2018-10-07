package net.cua.excel.util;

import org.dom4j.io.OutputFormat;
import org.dom4j.io.XMLWriter;

import java.io.IOException;
import java.io.OutputStream;
import java.io.UnsupportedEncodingException;

/**
 * Created by guanquan.wang on 2017/10/11.
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
     * @throws IOException
     *             DOCUMENT ME!
     */
    @Override
    protected void writeDeclaration() throws IOException {
        OutputFormat format = getOutputFormat();
        String encoding = format.getEncoding();

        // Only print of declaration is not suppressed
        if (!format.isSuppressDeclaration()) {
            // Assume 1.0 version
            if (encoding.equals("UTF8")) {
                writer.write("<?xml version=\"1.0\"");

                if (!format.isOmitEncoding()) {
                    writer.write(" encoding=\"UTF-8\"");
                }

                writer.write(" standalone=\"yes\"");

                writer.write("?>");
            } else {
                writer.write("<?xml version=\"1.0\"");

                if (!format.isOmitEncoding()) {
                    writer.write(" encoding=\"" + encoding + "\"");
                }

                writer.write(" standalone=\"yes\"");

                writer.write("?>");
            }

            if (format.isNewLineAfterDeclaration()) {
                println();
            }
        }
    }
}

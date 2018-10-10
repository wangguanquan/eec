package net.cua.excel.reader;

import org.xml.sax.Attributes;
import org.xml.sax.SAXException;
import org.xml.sax.helpers.DefaultHandler;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.parsers.SAXParser;
import javax.xml.parsers.SAXParserFactory;
import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;

/**
 * Create by guanquan.wang at 2018-09-27 14:28
 */
class SharedString {
    private Path sstPath;

    SharedString(Path sstPath) {
        this.sstPath = sstPath;
    }

    SharedString load() throws IOException {
        // get unique count
        max_size = uniqueCount();
        sharedString = max_size < 0 || max_size > page_size ? new String[page_size] : new String[max_size];
        return this;
    }

    String get(int index) {
        if (index < 0 || max_size > -1 && max_size < index) {
            throw new IndexOutOfBoundsException("Index: "+index+", Size: "+max_size);
        }
        if (!arrayRange(index)) {
            // reload data
            offset = index / page_size * page_size;
            sharedString[0] = null; // start mark
            loadXml();
            if (sharedString[0] == null) {
                throw new IndexOutOfBoundsException("Index: "+index+", Size: "+max_size);
            }
        }
        return sharedString[index - offset];
    }

    protected boolean arrayRange(int index) {
        return offset >= 0 && (max_size < page_size || (index - offset) >= 0 && (index - offset) < page_size);
    }

    String[] sharedString;
    int page_size = 2048
            , max_size = -1 // unknown size
            , offset = -1; // offset of all word

    private int uniqueCount() throws IOException {
        try (InputStream is = Files.newInputStream(sstPath)) {
            byte[] bytes = new byte[512];
            int n = is.read(bytes);
            String line = new String(bytes, 0, n, StandardCharsets.UTF_8);
            String uniqueCount = " uniqueCount=";
            int index = line.indexOf(uniqueCount), end = index > 0 ? line.indexOf('"', index+=(uniqueCount.length()+1)) : -1;
            if (end > 0) {
                return Integer.parseInt(line.substring(index, end));
            }
        }
        return -1;
    }

    private void loadXml() {
        SAXParserFactory factory = SAXParserFactory.newInstance();
        try (InputStream is = Files.newInputStream(sstPath)) {
            SAXParser parser = factory.newSAXParser();
            parser.parse(is, new DefaultHandler() {
                boolean open = false;
                int index = 0, current = index, start = offset, end = start + page_size;
                StringBuilder buf = new StringBuilder();
                @Override
                public void startElement(String uri, String localName, String qName, Attributes attributes) throws SAXException {
                    if (index >= start && index < end) open = 't' == qName.charAt(0);
                }

                @Override
                public void characters(char[] ac, int start, int length) throws SAXException {
                    if (open) buf.append(ac, start, length); // append to cache
                }

                @Override
                public void endElement(String uri, String localName, String qName) throws SAXException {
                    if ('t' == qName.charAt(0)) {
                        open = false;
                        sharedString[current++] = buf.toString();
                        buf.delete(0, buf.length()); // clear cache
                        index++;
                        if (index > end && max_size > -1) { // break parser
                            throw new BreakParserException();
                        }
                    }
                }

                @Override
                public void endDocument () throws SAXException {
                    max_size = index;
                }
            });
        } catch (BreakParserException e) {
            // break parser xml
        } catch (ParserConfigurationException | SAXException | IOException e) {
            throw new ExcelReadException("Read SharedString error.");
        }
    }

    static class BreakParserException extends RuntimeException { }
}

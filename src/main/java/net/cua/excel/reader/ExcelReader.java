package net.cua.excel.reader;

import net.cua.excel.entity.e7.Relationship;
import net.cua.excel.manager.RelManager;
import net.cua.excel.util.FileUtil;
import net.cua.excel.util.ZipUtil;
import org.dom4j.*;
import org.dom4j.io.SAXReader;

import java.io.IOException;
import java.io.InputStream;
import java.io.UncheckedIOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.*;
import java.util.stream.Stream;
import java.util.stream.StreamSupport;

/**
 * Excel读取工具
 * Create by guanquan.wang at 2018-09-22
 */
public class ExcelReader implements AutoCloseable {
    private ExcelReader() {}
    private Path self;

    private Sheet[] sheets;

    public static ExcelReader read(Path path) throws IOException {
        return read(Files.newInputStream(path));
    }

    public static ExcelReader read(InputStream stream) throws IOException {
        // Store template stream as zip file
        Path temp = FileUtil.mktmp("eec+");
        ZipUtil.unzip(stream, temp);

        // load workbook.xml
        SAXReader reader = new SAXReader();
        Document document;
        try {
            document = reader.read(Files.newInputStream(temp.resolve("xl/_rels/workbook.xml.rels")));
        } catch (DocumentException | IOException e) {
            FileUtil.rm_rf(temp.toFile(), true);
            throw new ExcelReadException(e);
        }
        List<Element> list = document.getRootElement().elements();
        Relationship[] rels = new Relationship[list.size()];
        int i = 0;
        for (Element e : list) {
            rels[i++] = new Relationship(e.attributeValue("Id"), e.attributeValue("Target"), e.attributeValue("Type"));
        }
        RelManager relManager = RelManager.of(rels);

        try {
            document = reader.read(Files.newInputStream(temp.resolve("xl/workbook.xml")));
        } catch (DocumentException | IOException e) {
            // read style file fail.
            FileUtil.rm_rf(temp.toFile(), true);
            throw new ExcelReadException(e);
        }
        Element root = document.getRootElement();
        Namespace ns = root.getNamespaceForPrefix("r");

        // Load SharedString
        SharedString sst = new SharedString(temp.resolve("xl/sharedStrings.xml")).load();

        List<Sheet> sheets = new ArrayList<>();
        Iterator<Element> sheetIter = root.element("sheets").elementIterator();
        for (; sheetIter.hasNext(); ) {
            Element e = sheetIter.next();
            Sheet sheet = new Sheet();
            sheet.setName(e.attributeValue("name"));
            sheet.setIndex(Integer.parseInt(e.attributeValue("sheetId")));
            Relationship r = relManager.getById(e.attributeValue(QName.get("id", ns)));
            if (r == null) {
                FileUtil.rm_rf(temp.toFile(), true);
                throw new ExcelReadException("File has be destroyed");
            }
            sheet.setPath(temp.resolve("xl").resolve(r.getTarget()));
            // put shared string
            sheet.setSst(sst);
            sheets.add(sheet);
        }

        // sort by sheet index
        sheets.sort(Comparator.comparingInt(Sheet::getIndex));

        ExcelReader er = new ExcelReader();
        er.sheets = sheets.toArray(new Sheet[sheets.size()]);
        er.self = temp;

        return er;
    }

    /**
     * to streams
     * @return sheet stream
     */
    public Stream<Sheet> sheets() {
        Iterator<Sheet> iter = new Iterator<Sheet>() {
            int n = 0;
            @Override
            public boolean hasNext() {
                return n < sheets.length;
            }

            @Override
            public Sheet next() {
                try {
                    // test and load sheet data
                    return sheets[n++].load();
                } catch (IOException e) {
                    throw new UncheckedIOException(e);
                }
            }
        };
        return StreamSupport.stream(Spliterators.spliterator(iter, sheets.length
                , Spliterator.ORDERED | Spliterator.NONNULL), false);
    }

    /**
     * get by index
     * @param index
     * @return
     */
    public Sheet sheet(int index) {
        try {
            return sheets[index].load();
        } catch (IOException e) {
            throw new UncheckedIOException(e);
        }
    }

    /**
     * get by name
     * @param sheetName
     * @return
     */
    public Sheet sheet(String sheetName) {
        try {
            for (Sheet t : sheets) {
                if (sheetName.equals(t.getName())) {
                    return t.load();
                }
            }
        } catch (IOException e) {
            throw new UncheckedIOException(e);
        }
        return null;
    }

    /**
     *
     * @return size of sheets
     */
    public int getSize() {
        return sheets != null ? sheets.length : 0;
    }

    /**
     * close stream and delete temp files
     * @throws IOException
     */
    public void close() throws IOException {
        // close sheet
        for (Sheet st : sheets) {
            st.close();
        }
        // delete temp files
        FileUtil.rm_rf(self.toFile(), true);
    }
}

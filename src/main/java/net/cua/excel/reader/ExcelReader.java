package net.cua.excel.reader;

import net.cua.excel.entity.e7.Relationship;
import net.cua.excel.manager.ExcelType;
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
 * 一个流式操作链，使用游标控制，游标只会向前，所以不能反复操作同一个Sheet流。
 * 同一个Sheet页内部Row对象是内存共享的，所以不要直接将Stream<Row>转为集合类.
 * 你首先应该考虑使用try-with-resource使用Reader或手动关闭ExcelReader。
 * <code>
 *     try (ExcelReader reader = ExcelReader.read(path)) {
 *         reader.sheets().flatMap(Sheet::rows).forEach(System.out::println);
 *     } catch (IOException e) {}
 * </code>
 * Create by guanquan.wang at 2018-09-22
 */
public class ExcelReader implements AutoCloseable {
    private ExcelReader() {}
    private Path self;

    private Sheet[] sheets;

    /**
     * 实例化Reader
     * @param path Excel路径
     * @return ExcelReader
     * @throws IOException 文件不存在或读取文件失败
     */
    public static ExcelReader read(Path path) throws IOException {
        return read(Files.newInputStream(path), 0, 0);
    }
    /**
     * 实例化Reader
     * @param stream Excel文件流
     * @return ExcelReader
     * @throws IOException 读取文件失败
     */
    public static ExcelReader read(InputStream stream) throws IOException {
        return read(stream, 0, 0);
    }
    /**
     * 实例化Reader
     * @param path Excel路径
     * @param cacheSize sharedString缓存大小，默认512
     *                  将此参数影响读取文件次数
     * @return ExcelReader
     * @throws IOException 文件不存在或读取文件失败
     */
    public static ExcelReader read(Path path, int cacheSize) throws IOException {
        return read(Files.newInputStream(path), cacheSize, 0);
    }

    /**
     * 实例化Reader
     * @param stream Excel文件流
     * @param cacheSize sharedString缓存大小，默认512
     *                  将此参数影响读取文件次数
     * @return ExcelReader
     * @throws IOException 读取文件失败
     */
    public static ExcelReader read(InputStream stream, int cacheSize) throws IOException {
        return read(stream, cacheSize, 0);
    }
    /**
     * 实例化Reader
     * @param path Excel路径
     * @param cacheSize sharedString缓存大小，默认512
     *                  将此参数影响读取文件次数
     * @param hotSize 热词区大小，默认64
     * @return ExcelReader
     * @throws IOException 文件不存在或读取文件失败
     */
    public static ExcelReader read(Path path, int cacheSize, int hotSize) throws IOException {
        return read(Files.newInputStream(path), cacheSize, hotSize);
    }
    /**
     * 实例化Reader
     * @param stream Excel文件流
     * @param cacheSize sharedString缓存大小，默认512
     *                  将此参数影响读取文件次数
     * @param hotSize 热词区大小，默认64
     * @return ExcelReader
     * @throws IOException 读取文件失败
     */
    public static ExcelReader read(InputStream stream, int cacheSize, int hotSize) throws IOException {
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
        SharedString sst = new SharedString(temp.resolve("xl/sharedStrings.xml"), cacheSize, hotSize).load();

        List<Sheet> sheets = new ArrayList<>();
        Iterator<Element> sheetIter = root.element("sheets").elementIterator();
        for (; sheetIter.hasNext(); ) {
            Element e = sheetIter.next();
            Sheet sheet = new Sheet();
            sheet.setName(e.attributeValue("name"));
            sheet.setIndex(Integer.parseInt(e.attributeValue("sheetId")));
            String state = e.attributeValue("state");
            sheet.setHidden("hidden".equals(state));
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
     * @return sheet流
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
     * @param index sheet index of workbook
     * @return sheet
     */
    public Sheet sheet(int index) {
        try {
            return sheets[index].load(); // lazy loading worksheet data
        } catch (IOException e) {
            throw new UncheckedIOException(e);
        }
    }

    /**
     * get by name
     * @param sheetName name
     * @return null if not found
     */
    public Sheet sheet(String sheetName) {
        try {
            for (Sheet t : sheets) {
                if (sheetName.equals(t.getName())) {
                    return t.load(); // lazy loading worksheet data
                }
            }
        } catch (IOException e) {
            throw new UncheckedIOException(e);
        }
        return null;
    }

    /**
     * get all sheets
     * @return Sheet Array
     */
    public Sheet[] all() {
        return sheets;
    }

    /**
     * size of sheets
     * @return int
     */
    public int getSize() {
        return sheets != null ? sheets.length : 0;
    }

    /**
     * close stream and delete temp files
     * @throws IOException when fail close readers
     */
    public void close() throws IOException {
        // close sheet
        for (Sheet st : sheets) {
            st.close();
        }
        // delete temp files
        FileUtil.rm_rf(self.toFile(), true);
    }


    // --- check
    static ExcelType typeOfStream(byte[] bytes, int len) {
        ExcelType excelType = ExcelType.UNKNOWN;
        if (bytes.length < len || len < 4)
            return excelType;
        int type = bytes[0] & 0xff;
        type += (bytes[1] & 0xff) << 8;
        type += (bytes[2] & 0xff) << 16;
        type += (bytes[3] & 0xff) << 24;
        int zip = 0x04034b50;
        int biff1 = 0xe011cfd0;
        int biff2 = 0xe11ab1a1;

        if (type == zip) {
            excelType = ExcelType.XLSX;
        } else if (type == biff1 && len >= 8) {
            type = bytes[4] & 0xff;
            type += (bytes[5] & 0xff) << 8;
            type += (bytes[6] & 0xff) << 16;
            type += (bytes[7] & 0xff) << 24;
            if (type == biff2) excelType = ExcelType.BIFF8;
        }
        return excelType;
    }

}

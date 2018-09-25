package net.cua.export.reader;

import net.cua.export.entity.e7.SimpleTemplate;
import net.cua.export.util.FileUtil;
import net.cua.export.util.ZipUtil;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.sql.Connection;
import java.util.Spliterator;
import java.util.Spliterators;
import java.util.function.Predicate;
import java.util.stream.Stream;
import java.util.stream.StreamSupport;

/**
 * Create by guanquan.wang at 2018-09-22
 */
public class ExcelReader {
    private Connection con;

    private Sheet[] sheets;

    public static ExcelReader load(Path path) throws IOException {
        return load(Files.newInputStream(path));
    }

    public static ExcelReader load(InputStream stream) throws IOException {
        // Store template stream as zip file
        Path temp = FileUtil.mktmp("eec+");
        ZipUtil.unzip(stream, temp);


        return null;
    }

    public ExcelReader setConnection(Connection con) {
        this.con = con;
        return this;
    }

    public int importWithSQL(String sql) {
        return 0;
    }
    /**
     * to streams
     * @return sheet stream
     */
    public Stream<Sheet> sheets() {
        return StreamSupport.stream(Spliterators.spliterator(sheets
                , Spliterator.ORDERED | Spliterator.NONNULL), false);
    }

    /**
     *
     * @param index
     * @return
     */
    public Sheet sheet(int index) {
        return sheets[index];
    }

    public Sheet sheet(String sheetName) {
        for (Sheet t : sheets) {
            if (sheetName.equals(t.name)) {
                return t;
            }
        }
        return null;
    }

}

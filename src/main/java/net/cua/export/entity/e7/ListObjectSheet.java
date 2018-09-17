package net.cua.export.entity.e7;

import net.cua.export.annotation.DisplayName;
import net.cua.export.annotation.NotExport;
import net.cua.export.entity.ExportException;
import net.cua.export.entity.WaterMark;
import net.cua.export.util.ExtBufferedWriter;
import net.cua.export.util.StringUtil;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.lang.reflect.Field;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.sql.Timestamp;
import java.util.ArrayList;
import java.util.List;

/**
 * Created by guanquan.wang at 2018-01-26 14:48
 */
public class ListObjectSheet<T> extends Sheet {
    private List<T> data;

    public ListObjectSheet(Workbook workbook) {
        super(workbook);
    }

    public ListObjectSheet(Workbook workbook, String name, Column[] columns) {
        super(workbook, name, columns);
    }

    public ListObjectSheet(Workbook workbook, String name, WaterMark waterMark, Column[] columns) {
        super(workbook, name, waterMark, columns);
    }

    @Override
    public void close() {
        data.clear();
        data = null;
    }

    public ListObjectSheet<T> setData(final List<T> data) {
        this.data = data;
        return this;
    }

    @Override
    public void writeTo(Path xl) throws IOException, ExportException {
        Path worksheets = xl.resolve("worksheets");
        if (!Files.exists(worksheets)) {
            Files.createDirectory(worksheets);
        }
        String name = getFileName();
        logger.debug(getName() + " | " + name);

        // create header columns
        init();

        File sheetFile = worksheets.resolve(name).toFile();

        // write date
        try (ExtBufferedWriter bw = new ExtBufferedWriter(new OutputStreamWriter(new FileOutputStream(sheetFile), StandardCharsets.UTF_8))) {
            // Write header
            writeBefore(bw);
            // Main data
            // Write sheet data
            if (getAutoSize() == 1) {
                for (Object o : data) {
                    writeRowAutoSize(o, bw);
                }
            } else {
                for (Object o : data) {
                    writeRow(o, bw);
                }
            }

            // Write foot
            writeAfter(bw);

        } catch (IllegalAccessException e) {
            throw new ExportException(e);
        }

        // resize columns
        boolean resize = false;
        for (Column hc : columns) {
            if (hc.getWidth() > 0.000001) {
                resize = true;
                break;
            }
        }

        if (getAutoSize() == 1 || resize) {
            autoColumnSize(sheetFile);
        }

        // relationship
        relManager.write(worksheets, name);
    }

    private static final String[] exclude = {"serialVersionUID", "this$0"};
    protected void init() throws IOException {
        Object o = workbook.getFirst(data);
        if (columns == null || columns.length == 0) {
            Field[] fields = o.getClass().getDeclaredFields();
            List<Column> list = new ArrayList<>(fields.length);
            for (int i = 0; i < fields.length; i++) {
                Field field = fields[i];
                String gs = field.toGenericString();
                NotExport notExport = field.getAnnotation(NotExport.class);
                if (notExport != null || StringUtil.indexOf(exclude, gs.substring(gs.lastIndexOf('.') + 1)) >= 0) {
                    continue;
                }
                DisplayName dn = field.getAnnotation(DisplayName.class);
                if (dn != null && StringUtil.isNotEmpty(dn.value())) {
                    list.add(new Column(dn.value(), field.getName(), field.getType()).setShare(dn.share()));
                } else {
                    list.add(new Column(field.getName(), field.getName(), field.getType()).setShare(dn != null && dn.share()));
                }
            }
            columns = new Column[list.size()];
            list.toArray(columns);
            for (int i = 0; i < columns.length; i++) {
                columns[i].setSst(workbook.getStyles());
            }
        } else {
            for (Column hc : columns) {
                try {
                    if (hc.getClazz() == null) {
                        Field field = o.getClass().getDeclaredField(hc.getKey());
                        hc.setClazz(field.getType());
//                        DisplayName dn = field.getAnnotation(DisplayName.class);
//                        if (dn != null) {
//                            hc.setShare(hc.isShare() || dn.share());
//                            if (StringUtil.isEmpty(hc.getName())
//                                    && StringUtil.isNotEmpty(dn.value())) {
//                                hc.setName(dn.value());
//                            }
//                        }
                    }
                } catch (NoSuchFieldException e) {
                    throw new IOException(e.getCause());
                }
            }
        }

    }

    /**
     * write object
     *
     * @param o
     * @param bw
     * @throws IOException
     */
    protected void writeRowAutoSize(Object o, ExtBufferedWriter bw) throws IOException, IllegalAccessException {
        if (o == null) {
            writeEmptyRow(bw);
            return;
        }
        int r = ++rows;
        final int len = columns.length;
        bw.write("<row r=\"");
        bw.writeInt(r);
        bw.write("\" ht=\"16.5\" spans=\"1:");
        bw.writeInt(len);
        bw.write("\">");

        Field field;
        Object e;
        Class<?> oclazz = o.getClass();
        for (int i = 0; i < len; i++) {
            Column hc = columns[i];
            try {
                field = oclazz.getDeclaredField(hc.getKey());
                field.setAccessible(true);
            } catch (NoSuchFieldException e1) {
                throw new IOException(e1.getCause());
            }
            Class<?> clazz = hc.getClazz();
            // t n=numeric (default), s=string, b=boolean
            if (isString(clazz)) {
                String s = (e = field.get(o)) != null ? e.toString() : null;
                writeStringAutoSize(bw, s, i);
            }
            else if (isDate(clazz)) {
                e = field.get(o);
                if (e != null) {
                    java.sql.Date date = (java.sql.Date) e;
                    writeDate(bw, date, i);
                } else {
                    writeNull(bw, i);
                }
            }
            else if (isDateTime(clazz)) {
                e = field.get(o);
                if (e != null) {
                    Timestamp ts = (Timestamp) e;
                    writeTimestamp(bw, ts, i);
                } else {
                    writeNull(bw, i);
                }
            }
            else if (isChar(clazz)) {
                char c = field.getChar(o);
                writeCharAutoSize(bw, c, i);
            }
            else if (isInt(clazz)) {
                int n = field.getInt(o);
                writeIntAutoSize(bw, n, i);
            }
            else if (isLong(clazz)) {
                long l = field.getLong(o);
                writeLong(bw, l, i);
            }
            else if (isFloat(clazz)) {
                double d = field.getDouble(o);
                writeDouble(bw, d, i);
            }
            else if (isBool(clazz)) {
                boolean bool = field.getBoolean(o);
                writeBoolean(bw, bool, i);
            }
            else {
                String s = (e = field.get(o)) != null ? e.toString() : null;
                writeStringAutoSize(bw, s, i);
            }
        }
        bw.write("</row>");
    }

    protected void writeRow(Object o, ExtBufferedWriter bw) throws IOException, IllegalAccessException {
        if (o == null) {
            writeEmptyRow(bw);
            return;
        }
        // Row number
        int r = ++rows;
        final int len = columns.length;
        bw.write("<row r=\"");
        bw.writeInt(r);
        bw.write("\" ht=\"16.5\" spans=\"1:");
        bw.writeInt(len);
        bw.write("\">");

        Field field;
        Object e;
        Class<?> oclazz = o.getClass();
        for (int i = 0; i < len; i++) {
            Column hc = columns[i];
            try {
                field = oclazz.getDeclaredField(hc.getKey());
                field.setAccessible(true);
            } catch (NoSuchFieldException e1) {
                throw new IOException(e1.getCause());
            }
            Class<?> clazz = hc.getClazz();
            // t n=numeric (default), s=string, b=boolean
            if (isString(clazz)) {
                String s = (e = field.get(o)) != null ? e.toString() : null;
                writeString(bw, s, i);
            }
            else if (isDate(clazz)) {
                e = field.get(o);
                if (e != null) {
                    java.sql.Date date = (java.sql.Date) e;
                    writeDate(bw, date, i);
                } else {
                    writeNull(bw, i);
                }
            }
            else if (isDateTime(clazz)) {
                e = field.get(o);
                if (e != null) {
                    Timestamp ts = (Timestamp) e;
                    writeTimestamp(bw, ts, i);
                } else {
                    writeNull(bw, i);
                }
            }
            else if (isChar(clazz)) {
                char c = field.getChar(o);
                writeChar(bw, c, i);
            }
            else if (isInt(clazz)) {
                int n = field.getInt(o);
                writeInt(bw, n, i);
            }
            else if (isLong(clazz)) {
                long l = field.getLong(o);
                writeLong(bw, l, i);
            }
            else if (isFloat(clazz)) {
                double d = field.getDouble(o);
                writeDouble(bw, d, i);
            }
            else if (isBool(clazz)) {
                boolean bool = field.getBoolean(o);
                writeBoolean(bw, bool, i);
            }
            else {
                String s = (e = field.get(o)) != null ? e.toString() : null;
                writeString(bw, s, i);
            }
        }
        bw.write("</row>");
    }

}

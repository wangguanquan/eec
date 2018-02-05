package net.cua.export.entity.e7;

import net.cua.export.annotation.DisplayName;
import net.cua.export.annotation.NotExport;
import net.cua.export.entity.WaterMark;
import net.cua.export.util.ExtBufferedWriter;
import net.cua.export.util.StringUtil;
import org.apache.log4j.Logger;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.lang.reflect.Field;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.sql.Timestamp;
import java.util.ArrayList;
import java.util.List;

/**
 * Created by wanggq at 2018-01-26 14:48
 */
public class ListObjectSheet<T> extends Sheet {
    private Logger logger = Logger.getLogger(this.getClass().getName());
    private List<T> data;

    public ListObjectSheet(Workbook workbook) {
        super(workbook);
    }

    public ListObjectSheet(Workbook workbook, String name, HeadColumn[] headColumns) {
        super(workbook, name, headColumns);
    }

    public ListObjectSheet(Workbook workbook, String name, WaterMark waterMark, HeadColumn[] headColumns) {
        super(workbook, name, waterMark, headColumns);
    }

    public ListObjectSheet<T> setData(final List<T> data) {
        this.data = data;
        return this;
    }

    @Override
    public void writeTo(Path xl) throws IOException {
        Path worksheets = Paths.get(xl.toString(), "worksheets");
        if (!Files.exists(worksheets)) {
            Files.createDirectory(worksheets);
        }
        String name = getFileName();
//        logger.info(getName() + " | " + name);

        // create header columns
        init();

        File sheetFile = Paths.get(worksheets.toString(), name).toFile();

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

        } catch (IOException e) {
            e.printStackTrace();
        } catch (IllegalAccessException e) {
            e.printStackTrace();
        }

        // resize columns
        boolean resize = false;
        for (HeadColumn hc : headColumns) {
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
        if (headColumns == null || headColumns.length == 0) {
            Field[] fields = o.getClass().getDeclaredFields();
            List<HeadColumn> list = new ArrayList<>(fields.length);
            for (int i = 0; i < fields.length; i++) {
                Field field = fields[i];
                String gs = field.toGenericString();
                NotExport notExport = field.getAnnotation(NotExport.class);
                if (notExport != null || StringUtil.indexOf(exclude, gs.substring(gs.lastIndexOf('.') + 1)) >= 0) {
                    continue;
                }
                DisplayName dn = field.getAnnotation(DisplayName.class);
                if (dn != null && StringUtil.isNotEmpty(dn.value())) {
                    list.add(new HeadColumn(dn.value(), field.getName(), field.getType()).setShare(dn.share()));
                } else {
                    list.add(new HeadColumn(field.getName(), field.getName(), field.getType()).setShare(dn != null && dn.share()));
                }
            }
            headColumns = new HeadColumn[list.size()];
            list.toArray(headColumns);
        } else {
            for (HeadColumn hc : headColumns) {
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
        final int len = headColumns.length;
        bw.write("<row r=\"");
        bw.writeInt(r);
        bw.write("\" ht=\"16.5\" spans=\"1:");
        bw.writeInt(len);
        bw.write("\">");

        Field field;
        Object e;
        Class<?> clazz = o.getClass();
        for (int i = 0; i < len; i++) {
            HeadColumn hc = headColumns[i];
            try {
                field = clazz.getDeclaredField(hc.getKey());
                field.setAccessible(true);
            } catch (NoSuchFieldException e1) {
                throw new IOException(e1.getCause());
            }
            // t n=numeric (default), s=string, b=boolean
            if (hc.getClazz() == String.class) {
                String s = (e = field.get(o)) != null ? e.toString() : null;
                writeStringAutoSize(bw, s, i);
            }
            else if (hc.getClazz() == java.util.Date.class
                    || hc.getClazz() == java.sql.Date.class) {
                e = field.get(o);
                if (e != null) {
                    java.sql.Date date = (java.sql.Date) e;
                    writeDate(bw, date, i);
                } else {
                    writeNull(bw, i);
                }
            }
            else if (hc.getClazz() == java.sql.Timestamp.class) {
                e = field.get(o);
                if (e != null) {
                    Timestamp ts = (Timestamp) e;
                    writeTimestamp(bw, ts, i);
                } else {
                    writeNull(bw, i);
                }
            }
            else if (hc.getClazz() == int.class || hc.getClazz() == Integer.class
                    || hc.getClazz() == byte.class || hc.getClazz() == Byte.class
                    || hc.getClazz() == short.class || hc.getClazz() == Short.class
                    ) {
                int n = field.getInt(o);
                writeIntAutoSize(bw, n, i);
            }
            else if (hc.getClazz() == char.class || hc.getClazz() == Character.class) {
                char c = field.getChar(o);
                writeCharAutoSize(bw, c, i);
            }
            else if (hc.getClazz() == long.class || hc.getClazz() == Long.class) {
                long l = field.getLong(o);
                writeLong(bw, l, i);
            }
            else if (hc.getClazz() == double.class || hc.getClazz() == Double.class
                    || hc.getClazz() == float.class || hc.getClazz() == Float.class
                    ) {
                double d = field.getDouble(o);
                writeDouble(bw, d, i);
            }
            else if (hc.getClazz() == boolean.class || hc.getClazz() == Boolean.class) {
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
        final int len = headColumns.length;
        bw.write("<row r=\"");
        bw.writeInt(r);
        bw.write("\" ht=\"16.5\" spans=\"1:");
        bw.writeInt(len);
        bw.write("\">");

        Field field;
        Object e;
        Class<?> clazz = o.getClass();
        for (int i = 0; i < len; i++) {
            HeadColumn hc = headColumns[i];
            try {
                field = clazz.getDeclaredField(hc.getKey());
                field.setAccessible(true);
            } catch (NoSuchFieldException e1) {
                throw new IOException(e1.getCause());
            }
            // t n=numeric (default), s=string, b=boolean
            if (hc.getClazz() == String.class) {
                String s = (e = field.get(o)) != null ? e.toString() : null;
                writeString(bw, s, i);
            }
            else if (hc.getClazz() == java.util.Date.class
                    || hc.getClazz() == java.sql.Date.class) {
                e = field.get(o);
                if (e != null) {
                    java.sql.Date date = (java.sql.Date) e;
                    writeDate(bw, date, i);
                } else {
                    writeNull(bw, i);
                }
            }
            else if (hc.getClazz() == java.sql.Timestamp.class) {
                e = field.get(o);
                if (e != null) {
                    Timestamp ts = (Timestamp) e;
                    writeTimestamp(bw, ts, i);
                } else {
                    writeNull(bw, i);
                }
            }
            else if (hc.getClazz() == int.class || hc.getClazz() == Integer.class
                    || hc.getClazz() == byte.class || hc.getClazz() == Byte.class
                    || hc.getClazz() == short.class || hc.getClazz() == Short.class
                    ) {
                int n = field.getInt(o);
                writeInt(bw, n, i);
            }
            else if (hc.getClazz() == char.class || hc.getClazz() == Character.class) {
                char c = field.getChar(o);
                writeChar(bw, c, i);
            }
            else if (hc.getClazz() == long.class || hc.getClazz() == Long.class) {
                long l = field.getLong(o);
                writeLong(bw, l, i);
            }
            else if (hc.getClazz() == double.class || hc.getClazz() == Double.class
                    || hc.getClazz() == float.class || hc.getClazz() == Float.class
                    ) {
                double d = field.getDouble(o);
                writeDouble(bw, d, i);
            }
            else if (hc.getClazz() == boolean.class || hc.getClazz() == Boolean.class) {
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

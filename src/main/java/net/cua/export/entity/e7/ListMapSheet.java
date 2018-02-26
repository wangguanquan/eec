package net.cua.export.entity.e7;

import net.cua.export.entity.WaterMark;
import net.cua.export.util.ExtBufferedWriter;
import net.cua.export.util.StringUtil;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.sql.Date;
import java.sql.Timestamp;
import java.util.List;
import java.util.Map;

/**
 * Created by guanquan.wang at 2018-01-26 14:46
 */
public class ListMapSheet extends Sheet {
    List<Map<String, ?>> data;

    public ListMapSheet(Workbook workbook) {
        super(workbook);
    }

    public ListMapSheet(Workbook workbook, String name, HeadColumn[] headColumns) {
        super(workbook, name, headColumns);
    }

    public ListMapSheet(Workbook workbook, String name, WaterMark waterMark, HeadColumn[] headColumns) {
        super(workbook, name, waterMark, headColumns);
    }

    @Override
    public void close() {
        data.clear();
        data = null;
    }

    public ListMapSheet setData(final List<Map<String, ?>> data) {
        this.data = data;
        return this;
    }

    @Override
    public void writeTo(Path xl) throws IOException {
        Path worksheets = xl.resolve("worksheets");
        if (!Files.exists(worksheets)) {
            Files.createDirectory(worksheets);
        }
        String name = getFileName();
//        logger.info(getName() + " | " + name);

        Map<String, ?> first = (Map<String, ?>) workbook.getFirst(data);
        for (int i = 0; i < headColumns.length; i++) {
            HeadColumn hc = headColumns[i];
            if (StringUtil.isEmpty(hc.getKey())) {
                throw new IOException(getClass() + " 类别必须指定map的key。");
            }
            if (hc.getClazz() == null) {
                hc.setClazz(first.get(hc.getKey()).getClass());
            }
        }

        File sheetFile = worksheets.resolve(name).toFile();

        // write date
        try (ExtBufferedWriter bw = new ExtBufferedWriter(new OutputStreamWriter(new FileOutputStream(sheetFile), StandardCharsets.UTF_8))) {
            // Write header
            writeBefore(bw);
            // Main data
            // Write sheet data
            if (getAutoSize() == 1) {
                for (Map<String, ?> map : data) {
                    writeRowAutoSize(map, bw);
                }
            } else {
                for (Map<String, ?> map : data) {
                    writeRow(map, bw);
                }
            }

            // Write foot
            writeAfter(bw);

        } catch (IOException e) {
            throw e;
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

    /**
     * write map
     *
     * @param map
     * @param bw
     * @throws IOException
     */
    protected void writeRowAutoSize(Map<String, ?> map, ExtBufferedWriter bw) throws IOException {
        if (map == null) {
            writeEmptyRow(bw);
            return;
        }
        // 行番号
        int r = ++rows;
        final int len = headColumns.length;
        bw.write("<row r=\"");
        bw.writeInt(r);
        bw.write("\" ht=\"16.5\" spans=\"1:");
        bw.writeInt(len);
        bw.write("\">");

        for (int i = 0; i < len; i++) {
            HeadColumn hc = headColumns[i];
            Object o = map.get(hc.getKey());
            if (o != null) {
                // t n=numeric (default), s=string, b=boolean
                if (hc.getClazz() == String.class) {
                    String s = o.toString();
                    writeStringAutoSize(bw, s, i);
                }
                else if (hc.getClazz() == java.util.Date.class
                        || hc.getClazz() == java.sql.Date.class) {
                    java.sql.Date date = (Date) o;
                    writeDate(bw, date, i);
                }
                else if (hc.getClazz() == java.sql.Timestamp.class) {
                    Timestamp ts = (Timestamp) o;
                    writeTimestamp(bw, ts, i);
                }
                else if (hc.getClazz() == int.class || hc.getClazz() == Integer.class
                        || hc.getClazz() == byte.class || hc.getClazz() == Byte.class
                        || hc.getClazz() == short.class || hc.getClazz() == Short.class
                        ) {
                    int n = ((Integer) o).intValue();
                    writeIntAutoSize(bw, n, i);
                }
                else if (hc.getClazz() == char.class || hc.getClazz() == Character.class) {
                    char c = ((Character) o).charValue();
                    writeCharAutoSize(bw, c, i);
                }
                else if (hc.getClazz() == long.class || hc.getClazz() == Long.class) {
                    long l = ((Long) o).longValue();
                    writeLong(bw, l, i);
                }
                else if (hc.getClazz() == double.class || hc.getClazz() == Double.class
                        || hc.getClazz() == float.class || hc.getClazz() == Float.class
                        ) {
                    double d = ((Double) o).doubleValue();
                    writeDouble(bw, d, i);
                }
                else if (hc.getClazz() == boolean.class || hc.getClazz() == Boolean.class) {
                    boolean bool = ((Boolean) o).booleanValue();
                    writeBoolean(bw, bool, i);
                }
                else {
                    writeStringAutoSize(bw, o.toString(), i);
                }
            } else {
                writeNull(bw, i);
            }
        }
        bw.write("</row>");
    }

    protected void writeRow(Map<String, ?> map, ExtBufferedWriter bw) throws IOException {
        if (map == null) {
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

        for (int i = 0; i < len; i++) {
            HeadColumn hc = headColumns[i];
            Object o = map.get(hc.getKey());
            if (o != null) {
                // t n=numeric (default), s=string, b=boolean
                if (hc.getClazz() == String.class) {
                    String s = o.toString();
                    writeString(bw, s, i);
                }
                else if (hc.getClazz() == java.util.Date.class
                        || hc.getClazz() == java.sql.Date.class) {
                    java.sql.Date date = (Date) o;
                    writeDate(bw, date, i);
                }
                else if (hc.getClazz() == java.sql.Timestamp.class) {
                    Timestamp ts = (Timestamp) o;
                    writeTimestamp(bw, ts, i);
                }
                else if (hc.getClazz() == int.class || hc.getClazz() == Integer.class
                        || hc.getClazz() == byte.class || hc.getClazz() == Byte.class
                        || hc.getClazz() == short.class || hc.getClazz() == Short.class
                        ) {
                    int n = ((Integer) o).intValue();
                    writeInt(bw, n, i);
                }
                else if (hc.getClazz() == char.class || hc.getClazz() == Character.class) {
                    char c = ((Character) o).charValue();
                    writeChar(bw, c, i);
                }
                else if (hc.getClazz() == long.class || hc.getClazz() == Long.class) {
                    long l = ((Long) o).longValue();
                    writeLong(bw, l, i);
                }
                else if (hc.getClazz() == double.class || hc.getClazz() == Double.class
                        || hc.getClazz() == float.class || hc.getClazz() == Float.class
                        ) {
                    double d = ((Double) o).doubleValue();
                    writeDouble(bw, d, i);
                }
                else if (hc.getClazz() == boolean.class || hc.getClazz() == Boolean.class) {
                    boolean bool = ((Boolean) o).booleanValue();
                    writeBoolean(bw, bool, i);
                }
                else {
                    writeString(bw, o.toString(), i);
                }
            }
            else {
                writeNull(bw, i);
            }
        }
        bw.write("</row>");
    }
}

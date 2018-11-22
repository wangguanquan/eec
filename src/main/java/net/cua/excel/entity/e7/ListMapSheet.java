package net.cua.excel.entity.e7;

import net.cua.excel.util.ExtBufferedWriter;
import net.cua.excel.util.FileUtil;
import net.cua.excel.util.StringUtil;
import net.cua.excel.entity.WaterMark;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.math.BigDecimal;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.sql.Date;
import java.sql.Timestamp;
import java.time.LocalDate;
import java.time.LocalDateTime;
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

    public ListMapSheet(Workbook workbook, String name, Column[] columns) {
        super(workbook, name, columns);
    }

    public ListMapSheet(Workbook workbook, String name, WaterMark waterMark, Column[] columns) {
        super(workbook, name, waterMark, columns);
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
            FileUtil.mkdir(worksheets);
        }
        String name = getFileName();
        workbook.what("0010", getName());

        Map<String, ?> first = (Map<String, ?>) workbook.getFirst(data);
        for (int i = 0; i < columns.length; i++) {
            Column hc = columns[i];
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
        final int len = columns.length;
        bw.write("<row r=\"");
        bw.writeInt(r);
        bw.write("\" spans=\"1:");
        bw.writeInt(len);
        bw.write("\">");

        for (int i = 0; i < len; i++) {
            Column hc = columns[i];
            Object o = map.get(hc.getKey());
            if (o != null) {
                Class<?> clazz = hc.getClazz();
                // t n=numeric (default), s=string, b=boolean
                if (isString(clazz)) {
                    String s = o.toString();
                    writeStringAutoSize(bw, s, i);
                }
                else if (isDate(clazz)) {
                    java.util.Date date = (java.util.Date) o;
                    writeDate(bw, date, i);
                }
                else if (isDateTime(clazz)) {
                    Timestamp ts = (Timestamp) o;
                    writeTimestamp(bw, ts, i);
                }
                else if (isChar(clazz)) {
                    char c = ((Character) o).charValue();
                    writeCharAutoSize(bw, c, i);
                }
                else if (isInt(clazz)) {
                    int n = ((Integer) o).intValue();
                    writeIntAutoSize(bw, n, i);
                }
                else if (isLong(clazz)) {
                    long l = ((Long) o).longValue();
                    writeLong(bw, l, i);
                }
                else if (isFloat(clazz)) {
                    double d = ((Double) o).doubleValue();
                    writeDouble(bw, d, i);
                }
                else if (isBool(clazz)) {
                    boolean bool = ((Boolean) o).booleanValue();
                    writeBoolean(bw, bool, i);
                }
                else if (isBigDecimal(clazz)) {
                    writeBigDecimal(bw, (BigDecimal) o, i);
                }
                else if (isLocalDate(clazz)) {
                    writeLocalDate(bw, (LocalDate) o, i);
                }
                else if (isLocalDateTime(clazz)) {
                    writeLocalDateTime(bw, (LocalDateTime) o, i);
                }
                else if (isTime(clazz)) {
                    writeTime(bw, (java.sql.Time) o, i);
                }
                else if (isLocalTime(clazz)) {
                    writeLocalTime(bw, (java.time.LocalTime) o, i);
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
        final int len = columns.length;
        bw.write("<row r=\"");
        bw.writeInt(r);
        bw.write("\" spans=\"1:");
        bw.writeInt(len);
        bw.write("\">");

        for (int i = 0; i < len; i++) {
            Column hc = columns[i];
            Object o = map.get(hc.getKey());
            if (o != null) {
                Class<?> clazz = hc.getClazz();
                // t n=numeric (default), s=string, b=boolean
                if (isString(clazz)) {
                    String s = o.toString();
                    writeString(bw, s, i);
                }
                else if (isDate(clazz)) {
                    java.sql.Date date = (Date) o;
                    writeDate(bw, date, i);
                }
                else if (isDateTime(clazz)) {
                    Timestamp ts = (Timestamp) o;
                    writeTimestamp(bw, ts, i);
                }
                else if (isChar(clazz)) {
                    char c = ((Character) o).charValue();
                    writeChar(bw, c, i);
                }
                else if (isInt(clazz)) {
                    int n = ((Integer) o).intValue();
                    writeInt(bw, n, i);
                }
                else if (isLong(clazz)) {
                    long l = ((Long) o).longValue();
                    writeLong(bw, l, i);
                }
                else if (isFloat(clazz)) {
                    double d = ((Double) o).doubleValue();
                    writeDouble(bw, d, i);
                }
                else if (isBool(clazz)) {
                    boolean bool = ((Boolean) o).booleanValue();
                    writeBoolean(bw, bool, i);
                }
                else if (isBigDecimal(clazz)) {
                    writeBigDecimal(bw, (BigDecimal) o, i);
                }
                else if (isLocalDate(clazz)) {
                    writeLocalDate(bw, (LocalDate) o, i);
                }
                else if (isLocalDateTime(clazz)) {
                    writeLocalDateTime(bw, (LocalDateTime) o, i);
                }
                else if (isTime(clazz)) {
                    writeTime(bw, (java.sql.Time) o, i);
                }
                else if (isLocalTime(clazz)) {
                    writeLocalTime(bw, (java.time.LocalTime) o, i);
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

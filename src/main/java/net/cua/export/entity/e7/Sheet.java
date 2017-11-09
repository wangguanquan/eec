package net.cua.export.entity.e7;

import net.cua.export.annotation.TopNS;
import net.cua.export.entity.e7.Relationship;
import net.cua.export.entity.e7.SharedStrings;
import net.cua.export.entity.e7.Styles;
import net.cua.export.manager.Const;
import net.cua.export.manager.RelManager;
import net.cua.export.processor.ConversionStringProcessor;
import net.cua.export.processor.StyleProcessor;
import net.cua.export.util.ExtBufferedWriter;
import net.cua.export.util.StringUtil;
import org.apache.log4j.Logger;

import java.io.*;
import java.nio.Buffer;
import java.nio.ByteBuffer;
import java.nio.MappedByteBuffer;
import java.nio.channels.FileChannel;
import java.nio.charset.CharacterCodingException;
import java.nio.charset.StandardCharsets;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Timestamp;
import java.util.Arrays;
import java.util.concurrent.atomic.AtomicInteger;

import static net.cua.export.util.DateUtil.toDateTimeValue;
import static net.cua.export.util.DateUtil.toDateValue;

/**
 * Created by wanggq on 2017/9/26.
 */
@TopNS(prefix = {"", "r"}, value = "worksheet", uri = {Const.SCHEMA_MAIN, Const.Relationship.RELATIONSHIP})
public abstract class Sheet {
    private Logger logger = Logger.getLogger(this.getClass().getName());
    protected Workbook workbook;

    protected String name;
    protected HeadColumn[] headColumns;
    protected String waterMark;
    protected RelManager relManager;
    protected int id;
    private int autoSize;
    private double width = 20;
    private int baseInfoLen;
    protected int rows;

    public int getId() {
        return id;
    }

    public void setId(int id) {
        this.id = id;
    }

    public Sheet(Workbook workbook) {
        this.workbook = workbook;
        relManager = new RelManager();
    }
    public Sheet(Workbook workbook, String name, HeadColumn[] headColumns) {
        this.workbook = workbook;
        this.name = name;
        this.headColumns = headColumns;
        relManager = new RelManager();
    }

    public Sheet(Workbook workbook, String name, String waterMark, HeadColumn[] headColumns) {
        this.workbook = workbook;
        this.name = name;
        this.headColumns = headColumns;
        this.waterMark = waterMark;
        relManager = new RelManager();
    }

    public static class HeadColumn {
        private String name; // 列名
        private Class clazz;     // 列类型
        private boolean share; // 字符串是否共享
        private int type; // 0: 正常显示 1:显示百分比 2:显示人民币
        private ConversionStringProcessor processor;
        private StyleProcessor styleProcessor;
        private int cellStyle = -1; // 未设定
        private double width;
        private Object o;

        public HeadColumn() {}
        public HeadColumn(String name, Class clazz) {
            this(name, clazz, false);
        }

        public HeadColumn(String name, Class clazz, ConversionStringProcessor processor) {
            this(name, clazz, false);
            this.processor = processor;
        }

//        public HeadColumn(String name, Class clazz, StyleProcessor processor) {
//            this(name, clazz, false);
//            this.styleProcessor = processor;
//        }

        public HeadColumn(String name, Class clazz, boolean share) {
            this.name = name;
            this.clazz = clazz;
            this.share = share;
        }

        public HeadColumn(String name, Class clazz, ConversionStringProcessor processor, boolean share) {
            this.name = name;
            this.clazz = clazz;
            this.share = share;
            this.processor = processor;
        }

//        public HeadColumn(String name, Class clazz, StyleProcessor processor, boolean share) {
//            this.name = name;
//            this.clazz = clazz;
//            this.share = share;
//            this.styleProcessor = processor;
//        }

        public HeadColumn(String name, Class clazz, int cellStyle) {
            this.name = name;
            this.clazz = clazz;
            this.cellStyle = cellStyle;
        }

        public HeadColumn(String name, Class clazz, int cellStyle, boolean share) {
            this.name = name;
            this.clazz = clazz;
            this.cellStyle = cellStyle;
            this.share = share;
        }

        public HeadColumn setWidth(double width) {
            if (width < 0.00000001) {
                throw new RuntimeException("Width " + width + " less than 0.");
            }
            this.width = width;
            return this;
        }

        public boolean isShare() {
            return share;
        }

        public HeadColumn setType(int type) {
            this.type = type;
            return this;
        }

        public String getName() {
            return name;
        }

        public HeadColumn setName(String name) {
            this.name = name;
            return this;
        }

        public Class getClazz() {
            return clazz;
        }

        public HeadColumn setClazz(Class clazz) {
            this.clazz = clazz;
            return this;
        }

        public HeadColumn setProcessor(ConversionStringProcessor processor) {
            this.processor = processor;
            return this;
        }

        public HeadColumn setStyleProcessor(StyleProcessor styleProcessor) {
            this.styleProcessor = styleProcessor;
            return this;
        }

        public double getWidth() {
            return width;
        }

        public HeadColumn setCellStyle(int cellStyle) {
            this.cellStyle = cellStyle;
            return this;
        }

        protected int getCellStyle(Class clazz) {
            int style;
            if (clazz == String.class) {
                style = Styles.defaultStringStyle();
            } else if (clazz == java.sql.Date.class
                    || clazz == java.util.Date.class) {
                style = Styles.defaultDateStyle();
            } else if (clazz == java.sql.Timestamp.class) {
                style = Styles.defaultTimestampStyle();
            } else if (clazz == int.class || clazz == Integer.class
                    || clazz == long.class || clazz == Long.class
                    || clazz == char.class || clazz == Character.class
                    || clazz == byte.class || clazz == Byte.class
                    || clazz == short.class || clazz == Short.class
                    ) {
                style = Styles.defaultIntStyle();
                switch (type) {
                    case 0: // 正常显示数字
                        break;
                    case 1: // 百分比显示
                        style = Styles.clearNumfmt(style) | Styles.NumFmts.PADDING_PERCENTAGE_INT;
                        break;
                    case 2: // 显示人民币
                        style = Styles.clearNumfmt(style) | Styles.NumFmts.PADDING_YEN_INT;
                        break;
                    default:
                }
            } else if (clazz == double.class || clazz == Double.class
                    || clazz == float.class || clazz == Float.class
                    ) {
                style = Styles.defaultDoubleStyle();
                switch (type) {
                    case 0: // 正常显示数字
                    break;
                case 1: // 百分比显示
                    style= Styles.clearNumfmt(style) | Styles.NumFmts.PADDING_PERCENTAGE_DOUBLE;
                    break;
                case 2: // 显示人民币
                    style = Styles.clearNumfmt(style) | Styles.NumFmts.PADDING_YEN_DOUBLE;
                    break;
                default:
            }
            } else {
                style = 0;
            }
            return style;
        }
        public int getCellStyle() {
            if (cellStyle != -1) {
                return cellStyle;
            }
            return cellStyle = getCellStyle(clazz);
        }
    }

    public void autoSize() {
        this.autoSize = 1;
    }

    public void fixSize() {
        this.autoSize = 2;
    }

    public void fixSize(double width) {
        this.autoSize = 2;
        for (HeadColumn hc : headColumns) {
            hc.setWidth(width);
        }
    }

    public int getAutoSize() {
        return autoSize;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public final HeadColumn[] getHeadColumns() {
        return headColumns.clone();
    }

    public void setHeadColumns(final HeadColumn[] headColumns) {
        this.headColumns = headColumns.clone();
    }

    public String getWaterMark() {
        return waterMark;
    }

    public void setWaterMark(String waterMark) {
        this.waterMark = waterMark;
    }

    public void close() {
        for (HeadColumn o : headColumns) {
            o = null;
        }
        headColumns = null;
}

    public void addRel(Relationship rel) {
        relManager.add(rel);
    }

    public abstract void writeTo(File root);

    protected String getFileName() {
        return "sheet" + id + ".xml";
    }


    /**
     * 写worksheet头部
     * @param bw
     */
    protected void writeBefore(ExtBufferedWriter bw) throws IOException {
        // Declaration
        bw.write(Const.EXCEL_XML_DECLARATION);
        bw.newLine();
        // Root node
        if (getClass().isAnnotationPresent(TopNS.class)) {
            TopNS topNS = getClass().getAnnotation(TopNS.class);
            bw.write('<');
            bw.write(topNS.value());
            String[] prefixs = topNS.prefix(), urls = topNS.uri();
            for (int i = 0, len = prefixs.length; i < len; ) {
                bw.write(" xmlns");
                if (prefixs[i] != null && !prefixs[i].isEmpty()) {
                    bw.write(':');
                    bw.write(prefixs[i]);
                }
                bw.write("=\"");
                bw.write(urls[i]);
                if (++i < len) {
                    bw.write('"');
                }
            }
        } else {
            bw.write("<worksheet xmlns=\"");
            bw.write(Const.SCHEMA_MAIN);
        }
        bw.write("\">");

        // Dimension
        // Reset dimension at final
        bw.write("<dimension ref=\"A1\"/>");

        // SheetViews default value
        StringBuilder buf = new StringBuilder("<sheetViews><sheetView workbookViewId=\"0\"");
        if (id == 1) { // Default select the first worksheet
            buf.append(" tabSelected=\"1\"");
        }
//        <sheetView workbookViewId="0" tabSelected="1">
//            <selection sqref="C4" activeCell="C4"/>
//        </sheetView>
        buf.append("/></sheetViews>");

        // Default format
        buf.append("<sheetFormatPr defaultRowHeight=\"13.5\" baseColWidth=\"");
        buf.append((int)width);
        buf.append("\"/>");

        baseInfoLen = buf.length();
        // Write base info
        bw.write(buf.toString());

        // Write body data
        bw.write("<sheetData>");

        // Write header
        int r = ++rows;
        bw.write("<row r=\"");
        bw.writeInt(r);
        bw.write("\" customHeight=\"1\" ht=\"18.6\" spans=\"1:"); // spans 指定row开始和结束行
        bw.writeInt(headColumns.length);
        bw.write("\">");

        int c = 1, defaultStyle = defaultHeadStyle();
        for (HeadColumn hc : headColumns) {
            bw.write("<c r=\"");
            bw.write(int2Col(c++));
            bw.writeInt(r);
            bw.write("\" t=\"inlineStr\" s=\"");
            bw.writeInt(defaultStyle);
            bw.write("\"><is><t>");
            bw.write(hc.getName());
            bw.write("</t></is></c>");
        }
        bw.write("</row>");
    }

    int defaultHeadStyle() {
        return workbook.getStyles().of(
                Styles.Fonts.WHITE_GB2312_WRYH_11_B
                | Styles.Fills.FF666699
                | Styles.Borders.THIN_BLACK
                | Styles.Verticals.CENTER
                | Styles.Horizontals.CENTER);
    }

    /**
     * 写尾部
     * @param bw
     */
    protected void writeAfter(ExtBufferedWriter bw) throws IOException {
        SharedStrings sst = workbook.getSst();
        int c = 0;
        for (HeadColumn hc : headColumns) {
            if (hc.clazz == String.class && hc.isShare()) {
                c++;
            }
        }
        sst.addCount(c * (rows - 1));

        // End target --sheetData
        bw.write("</sheetData>");


        // TODO If autoSize ...


        // background image
        if (StringUtil.isNotEmpty(waterMark)) {
            // relationship
            Relationship r = relManager.likeByTarget("media/image"); // only one background image
            if (r != null) {
                bw.write("<picture r:id=\"");
                bw.write(r.getId());
                bw.write("\"/>");
            }
        }
        // End target
        if (getClass().isAnnotationPresent(TopNS.class)) {
            TopNS topNS = getClass().getAnnotation(TopNS.class);
            bw.write("</");
            bw.write(topNS.value());
            bw.write('>');
        } else {
            bw.write("</worksheet>");
        }
    }

    /**
     * 写行数据
     * @param rs ResultSet
     * @param bw
     */
    protected void writeRow(ResultSet rs, ExtBufferedWriter bw, SharedStrings sst, Styles styles) throws IOException, SQLException {
        // Row number
        int r = ++rows;
        final int len = headColumns.length;
        bw.write("<row r=\"");
        bw.writeInt(r);
        bw.write("\" ht=\"16.5\" spans=\"1:");
        bw.writeInt(len);
        bw.write("\">");

        for (int i = 1; i <= len; i++) {
            HeadColumn hc = headColumns[i - 1];
            bw.write("<c r=\"");
            bw.write(int2Col(i));
            bw.writeInt(r);

            int style = hc.cellStyle == -1 ? hc.getCellStyle() : hc.cellStyle;
            int styleIndex = styles.of(style);
            // t n=numeric (default), s=string, b=boolean
            if (hc.clazz == String.class) {
                String s = rs.getString(i);
                if (hc.styleProcessor != null) {
                    style = hc.styleProcessor.build(s, style);
                    styleIndex = styles.of(style);
                }
                if (StringUtil.isEmpty(s)) {
                    bw.write("\" s=\"");
                    bw.writeInt(styleIndex);
                    bw.write("\"/>");
                }
                else if (hc.isShare()) {
                    bw.write("\" t=\"s\" s=\"");
                    bw.writeInt(styleIndex);
                    bw.write("\"><v>");
                    bw.writeInt(sst.get(s));
                    bw.write("</v></c>");
                }
                else {
                    bw.write("\" t=\"inlineStr\" s=\"");
                    bw.writeInt(styleIndex);
                    bw.write("\"><is><t>");
                    bw.write(s);
                    bw.write("</t></is></c>");
                }
            }
            else if (hc.clazz == java.util.Date.class
                    || hc.clazz == java.sql.Date.class) {
                java.sql.Date date = rs.getDate(i);
                if (hc.styleProcessor != null) {
                    style = hc.styleProcessor.build(date, style);
                    styleIndex = styles.of(style);
                }
                if (date == null) {
                    bw.write("\" s=\"");
                    bw.writeInt(styleIndex);
                    bw.write("\"/>");
                } else {
                    bw.write("\" s=\"");
                    bw.writeInt(styleIndex);
                    bw.write("\"><v>");
                    bw.writeInt(toDateValue(date));
                    bw.write("</v></c>");
                }
            }
            else if (hc.clazz == java.sql.Timestamp.class) {
                Timestamp ts = rs.getTimestamp(i);
                if (hc.styleProcessor != null) {
                    style = hc.styleProcessor.build(ts, style);
                    styleIndex = styles.of(style);
                }
                if (ts == null) {
                    bw.write("\" s=\"");
                    bw.writeInt(styleIndex);
                    bw.write("\"/>");
                } else {
                    bw.write("\" s=\"");
                    bw.writeInt(styleIndex);
                    bw.write("\"><v>");
                    bw.write(toDateTimeValue(ts));
                    bw.write("</v></c>");
                }
            }
            else if (hc.clazz == int.class || hc.clazz == Integer.class
                    || hc.clazz == long.class || hc.clazz == Long.class
                    || hc.clazz == char.class || hc.clazz == Character.class
                    || hc.clazz == byte.class || hc.clazz == Byte.class
                    || hc.clazz == short.class || hc.clazz == Short.class
                    ) {
                int n = rs.getInt(i);
                if (hc.processor == null) {
                    if (hc.styleProcessor != null) {
                        style = hc.styleProcessor.build(n, style);
                        styleIndex = styles.of(style);
                    }
                    bw.write("\" s=\"");
                    bw.writeInt(styleIndex);
                    bw.write("\"><v>");
                    bw.writeInt(n);
                    bw.write("</v></c>");
                } else {
//                    if (hc.styleProcessor != null) {
//                        style = hc.styleProcessor.build(n, style);
//                        styleIndex = styles.of(style);
//                    }
                    Object o = hc.processor.conversion(n);
                    Class<?> clazz = o.getClass();
                    if (clazz == String.class) {
                        if (hc.cellStyle == Styles.defaultIntStyle()) {
                            style = hc.getCellStyle(String.class);
                            styleIndex = styles.of(style);
                        }
                        if (hc.styleProcessor != null) {
                            style = hc.styleProcessor.build(n, style);
                            styleIndex = styles.of(style);
                        }
                        String s = (String) o;
                        if (hc.isShare()) {
                            bw.write("\" t=\"s\" s=\"");
                            bw.writeInt(styleIndex);
                            bw.write("\"><v>");
                            bw.writeInt(sst.get(s));
                            bw.write("</v></c>");
                        } else {
                            bw.write("\" t=\"inlineStr\" s=\"");
                            bw.writeInt(styleIndex);
                            bw.write("\"><is><t>");
                            bw.write(s);
                            bw.write("</t></is></c>");
                        }
                    }
                    else if (clazz == int.class || clazz == Integer.class
                            || clazz == long.class || clazz == Long.class
                            || clazz == char.class || clazz == Character.class
                            || clazz == byte.class || clazz == Byte.class
                            || clazz == short.class || clazz == Short.class
                            ) {
                        bw.write("\" s=\"");
                        bw.writeInt(styleIndex);
                        bw.write("\"><v>");
                        bw.writeInt(n);
                        bw.write("</v></c>");
                    }
                    else if (clazz == java.util.Date.class
                            || clazz == java.sql.Date.class) {
                        if (hc.cellStyle == Styles.defaultIntStyle()) {
                            style = hc.getCellStyle(java.util.Date.class);
                            styleIndex = styles.of(style);
                        }
                        if (hc.styleProcessor != null) {
                            style = hc.styleProcessor.build(n, style);
                            styleIndex = styles.of(style);
                        }
                        bw.write("\" s=\"");
                        bw.writeInt(styleIndex);
                        bw.write("\"><v>");
                        bw.writeInt(toDateValue(rs.getDate(i)));
                        bw.write("</v></c>");
                    }
                    else if (clazz == java.sql.Timestamp.class) {
                        if (hc.cellStyle == Styles.defaultIntStyle()) {
                            style = hc.getCellStyle(java.sql.Timestamp.class);
                            styleIndex = styles.of(style);
                        }
                        if (hc.styleProcessor != null) {
                            style = hc.styleProcessor.build(n, style);
                            styleIndex = styles.of(style);
                        }
                        bw.write("\" s=\"");
                        bw.writeInt(styleIndex);
                        bw.write("\"><v>");
                        bw.write(toDateTimeValue(rs.getTimestamp(i)));
                        bw.write("</v></c>");
                    }
                    else if (hc.clazz == double.class || hc.clazz == Double.class
                            || hc.clazz == float.class || hc.clazz == Float.class
                            ) {
                        if (hc.cellStyle == Styles.defaultIntStyle()) {
                            style = hc.getCellStyle(double.class);
                            styleIndex = styles.of(style);
                        }
                        if (hc.styleProcessor == null) {
                            style = hc.styleProcessor.build(n, style);
                            styleIndex = styles.of(style);
                        }
                        bw.write("\" s=\"");
                        bw.writeInt(styleIndex);
                        bw.write("\"><v>");
                        bw.write(rs.getDouble(i));
                        bw.write("</v></c>");
                    }
                }
            }
            else if (hc.clazz == double.class || hc.clazz == Double.class
                    || hc.clazz == float.class || hc.clazz == Float.class
                    ) {
                double d = rs.getDouble(i);
                if (hc.styleProcessor != null) {
                    style = hc.styleProcessor.build(d, style);
                    styleIndex = styles.of(style);
                }
                bw.write("\" s=\"");
                bw.writeInt(styleIndex);
                bw.write("\"><v>");
                bw.write(d);
                bw.write("</v></c>");
            }
        }
        bw.write("</row>");
    }

    /**
     * 写行数据
     * @param rs ResultSet
     * @param bw
     */
    protected void writeRowAutoSize(ResultSet rs, ExtBufferedWriter bw, SharedStrings sst, Styles styles) throws IOException, SQLException {
        // 行番号
        int r = ++rows;
        final int len = headColumns.length;
        bw.write("<row r=\"");
        bw.writeInt(r);
        bw.write("\" ht=\"16.5\" spans=\"1:");
        bw.writeInt(len);
        bw.write("\">");

        for (int i = 1; i <= len; i++) {
            HeadColumn hc = headColumns[i - 1];
            bw.write("<c r=\"");
            bw.write(int2Col(i));
            bw.writeInt(r);

            int style = hc.cellStyle == -1 ? hc.getCellStyle() : hc.cellStyle;
            int styleIndex = styles.of(style);
            // t n=numeric (default), s=string, b=boolean
            if (hc.clazz == String.class) {
                String s = rs.getString(i);
                if (hc.styleProcessor != null) {
                    style = hc.styleProcessor.build(s, style);
                    styleIndex = styles.of(style);
                }
                if (StringUtil.isEmpty(s)) {
                    bw.write("\" s=\"");
                    bw.writeInt(styleIndex);
                    bw.write("\"/>");
                    continue;
                }
                int ln = s.getBytes("GB2312").length;
                if (hc.width == 0 && (hc.o == null || (int) hc.o < ln)) {
                    hc.o = ln;
                }
                if (hc.isShare()) {
                    bw.write("\" t=\"s\" s=\"");
                    bw.writeInt(styleIndex);
                    bw.write("\"><v>");
                    bw.writeInt(sst.get(s));
                    bw.write("</v></c>");
                } else {
                    bw.write("\" t=\"inlineStr\" s=\"");
                    bw.writeInt(styleIndex);
                    bw.write("\"><is><t>");
                    bw.write(s);
                    bw.write("</t></is></c>");
                }
            }
            else if (hc.clazz == java.util.Date.class
                    || hc.clazz == java.sql.Date.class) {
                java.sql.Date date = rs.getDate(i);
                if (hc.styleProcessor != null) {
                    style = hc.styleProcessor.build(date, style);
                    styleIndex = styles.of(style);
                }
                if (date == null) {
                    bw.write("\" s=\"");
                    bw.writeInt(styleIndex);
                    bw.write("\"/>");
                } else {
                    bw.write("\" s=\"");
                    bw.writeInt(styleIndex);
                    bw.write("\"><v>");
                    bw.writeInt(toDateValue(date));
                    bw.write("</v></c>");
                }
            }
            else if (hc.clazz == java.sql.Timestamp.class) {
                Timestamp ts = rs.getTimestamp(i);
                if (hc.styleProcessor != null) {
                    style = hc.styleProcessor.build(ts, style);
                    styleIndex = styles.of(style);
                }
                if (ts == null) {
                    bw.write("\" s=\"");
                    bw.writeInt(styleIndex);
                    bw.write("\"/>");
                } else {
                    bw.write("\" s=\"");
                    bw.writeInt(styleIndex);
                    bw.write("\"><v>");
                    bw.write(toDateTimeValue(ts));
                    bw.write("</v></c>");
                }
            }
            else if (hc.clazz == int.class || hc.clazz == Integer.class
                    || hc.clazz == long.class || hc.clazz == Long.class
                    || hc.clazz == char.class || hc.clazz == Character.class
                    || hc.clazz == byte.class || hc.clazz == Byte.class
                    || hc.clazz == short.class || hc.clazz == Short.class
                    ) {
                int n = rs.getInt(i);
                if (hc.processor == null) {
                    if (hc.styleProcessor != null) {
                        style = hc.styleProcessor.build(n, style);
                        styleIndex = styles.of(style);
                    }
                    bw.write("\" s=\"");
                    bw.writeInt(styleIndex);
                    bw.write("\"><v>");
                    bw.writeInt(n);
                    bw.write("</v></c>");
                } else {
//                    if (hc.styleProcessor != null) {
//                        style = hc.styleProcessor.build(n, style);
//                        styleIndex = styles.of(style);
//                    }
                    Object o = hc.processor.conversion(n);
                    Class<?> clazz = o.getClass();
                    if (clazz == String.class) {
                        logger.info(Arrays.toString(Styles.unpack(hc.cellStyle)) + " " + o);
                        if (hc.cellStyle == Styles.defaultIntStyle()) {
                            style = hc.getCellStyle(String.class);
                            styleIndex = styles.of(style);
                        }
                        if (hc.styleProcessor != null) {
                            style = hc.styleProcessor.build(n, style);
                            styleIndex = styles.of(style);
                        }
                        String s = (String) o;
                        if (hc.isShare()) {
                            bw.write("\" t=\"s\" s=\"");
                            bw.writeInt(styleIndex);
                            bw.write("\"><v>");
                            bw.writeInt(sst.get(s));
                            bw.write("</v></c>");
                        } else {
                            bw.write("\" t=\"inlineStr\" s=\"");
                            bw.writeInt(styleIndex);
                            bw.write("\"><is><t>");
                            bw.write(s);
                            bw.write("</t></is></c>");
                        }
                    }
                    else if (clazz == int.class || clazz == Integer.class
                            || clazz == long.class || clazz == Long.class
                            || clazz == char.class || clazz == Character.class
                            || clazz == byte.class || clazz == Byte.class
                            || clazz == short.class || clazz == Short.class
                            ) {
                        bw.write("\" s=\"");
                        bw.writeInt(styleIndex);
                        bw.write("\"><v>");
                        bw.writeInt(n);
                        bw.write("</v></c>");
                    }
                    else if (clazz == java.util.Date.class
                            || clazz == java.sql.Date.class) {
                        if (hc.cellStyle == Styles.defaultIntStyle()) {
                            style = hc.getCellStyle(java.util.Date.class);
                            styleIndex = styles.of(style);
                        }
                        if (hc.styleProcessor != null) {
                            style = hc.styleProcessor.build(n, style);
                            styleIndex = styles.of(style);
                        }
                        bw.write("\" s=\"");
                        bw.writeInt(styleIndex);
                        bw.write("\"><v>");
                        bw.writeInt(toDateValue(rs.getDate(i)));
                        bw.write("</v></c>");
                    }
                    else if (clazz == java.sql.Timestamp.class) {
                        if (hc.cellStyle == Styles.defaultIntStyle()) {
                            style = hc.getCellStyle(java.sql.Timestamp.class);
                            styleIndex = styles.of(style);
                        }
                        if (hc.styleProcessor != null) {
                            style = hc.styleProcessor.build(n, style);
                            styleIndex = styles.of(style);
                        }
                        bw.write("\" s=\"");
                        bw.writeInt(styleIndex);
                        bw.write("\"><v>");
                        bw.write(toDateTimeValue(rs.getTimestamp(i)));
                        bw.write("</v></c>");
                    }
                    else if (hc.clazz == double.class || hc.clazz == Double.class
                            || hc.clazz == float.class || hc.clazz == Float.class
                            ) {
                        if (hc.cellStyle == Styles.defaultIntStyle()) {
                            style = hc.getCellStyle(double.class);
                            styleIndex = styles.of(style);
                        }
                        if (hc.styleProcessor == null) {
                            style = hc.styleProcessor.build(n, style);
                            styleIndex = styles.of(style);
                        }
                        bw.write("\" s=\"");
                        bw.writeInt(styleIndex);
                        bw.write("\"><v>");
                        bw.write(rs.getDouble(i));
                        bw.write("</v></c>");
                    }
                }
            }
            else if (hc.clazz == double.class || hc.clazz == Double.class
                    || hc.clazz == float.class || hc.clazz == Float.class
                    ) {
                double v = rs.getDouble(i);
                if (hc.width == 0 && (hc.o == null || ((double)hc.o) < v)) {
                    hc.o = v;
                }
                if (hc.styleProcessor != null) {
                    style = hc.styleProcessor.build(v, style);
                    styleIndex = styles.of(style);
                }
                bw.write("\" s=\"");
                bw.writeInt(styleIndex);
                bw.write("\"><v>");
                bw.write(v);
                bw.write("</v></c>");
            }
        }
        bw.write("</row>");
    }

    protected  void autoColumnSize(File sheet) {
        // resize each column width ...
        File temp = new File(sheet.getParent(), sheet.getName() + ".temp");
        sheet.renameTo(temp);

        FileChannel inChannel = null, outChannel = null;
        FileInputStream fis = null;
        FileOutputStream fos = null;
        try {
            fis = new FileInputStream(temp);
            fos = new FileOutputStream(sheet);
            inChannel = fis.getChannel();
            outChannel = fos.getChannel();

//            int n = ExtBufferedWriter.stringSize(id);
            int dimensionIndex = 230, sheetViewLen = baseInfoLen + 3;
            inChannel.transferTo(0, dimensionIndex, outChannel);
            ByteBuffer buffer = ByteBuffer.allocate(sheetViewLen);
            inChannel.read(buffer, dimensionIndex);
            buffer.compact();
            if (buffer.get() == '"') {
                char[] chars = int2Col(headColumns.length);
                String s = ':' + new String(chars) + rows + "\"";
                outChannel.write(ByteBuffer.wrap(s.getBytes()));
            }
            outChannel.write(buffer);

            StringBuilder buf = new StringBuilder();
            buf.append("<cols>");
            int i = 0;
            for (HeadColumn hc : headColumns) {
                i++;
                buf.append("<col customWidth=\"1\" width=\"");
                if (hc.width > 0.0000001) {
                    buf.append(hc.width);
                    buf.append("\" max=\"");
                    buf.append(i);
                    buf.append("\" min=\"");
                    buf.append(i);
                    buf.append("\"/>");
                } else if (autoSize == 1) {
                    int _l = hc.name.getBytes("GB2312").length, len;
                    // TODO 根据字体字号计算文本宽度
                    if (hc.clazz == String.class) {
                        len = (int) hc.o;
//                        len = hc.o.toString().getBytes("GB2312").length;
                    }
                    else if (hc.clazz == java.util.Date.class
                            || hc.clazz == java.sql.Date.class) {
                        len = 10;
                    }
                    else if (hc.clazz == java.sql.Timestamp.class) {
                        len = 20;
                    }
                    else if (hc.clazz == int.class || hc.clazz == Integer.class
                            || hc.clazz == long.class || hc.clazz == Long.class
                            || hc.clazz == char.class || hc.clazz == Character.class
                            || hc.clazz == byte.class || hc.clazz == Byte.class
                            || hc.clazz == short.class || hc.clazz == Short.class
                            ) {
                        // TODO 根据numFmt计算字符宽度
                        len = hc.type > 0 ? 12 :  11;
                    }
                    else if (hc.clazz == double.class || hc.clazz == Double.class
                            || hc.clazz == float.class || hc.clazz == Float.class
                            ) {
                        // TODO 根据numFmt计算字符宽度
                        len = hc.o.toString().getBytes("GB2312").length;
//                        if (len < 11) {
//                            len = hc.type > 0 ? 12 : 11;
//                        }
                    } else {
                        len = 10;
                    }
                    buf.append(_l > len ? _l + 3.38 : len + 3.38);
                    buf.append("\" max=\"");
                    buf.append(i);
                    buf.append("\" min=\"");
                    buf.append(i);
                    buf.append("\" bestFit=\"1\"/>");
                } else {
                    buf.append(width);
                    buf.append("\" max=\"");
                    buf.append(i);
                    buf.append("\" min=\"");
                    buf.append(i);
                    buf.append("\"/>");
                }
            }
            buf.append("</cols>");

            outChannel.write(ByteBuffer.wrap(buf.toString().getBytes()));
            int start = dimensionIndex + sheetViewLen;
            inChannel.transferTo(start, inChannel.size() - start, outChannel);

        } catch (IOException e) {
            logger.error(e);
        } finally {
            if (fis != null) {
                try {
                    fis.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
            if (fos != null) {
                try {
                    fos.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
            if (inChannel != null) {
                try {
                    inChannel.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
            if (outChannel != null) {
                try {
                    outChannel.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
            boolean delete = temp.delete();
            if (!delete) {
                logger.error("Delete temp file failed.");
            }
        }

    }

    private final char[][] cache_col = {{'A'}, {'A', 'A'}, {'A', 'A', 'A'}};
    protected char[] int2Col(int n) {
        char[] c;
        if (n <= 26) {
            c = cache_col[0];
            c[0] = (char) (n - 1 + 'A');
        } else if (n <= 702) {
            int t = n / 26, w = n % 26;
            if (w == 0) {
                t--;
                w = 26;
            }
            c = cache_col[1];
            c[0] = (char) (t - 1 + 'A');
            c[1] = (char) (w - 1 + 'A');
        } else {
            int t = n / 26, tt = t / 26, w = n % 26;
            if (w == 0) {
                t--;
                w = 26;
            }
            c = cache_col[2];
            c[0] = (char) (tt - 1 + 'A');
            c[1] = (char) (t - 27 + 'A');
            c[2] = (char) (w - 1 + 'A');
        }
        return c;
    }

}

package net.cua.export.entity.e7;

import net.cua.export.annotation.TopNS;
import net.cua.export.entity.ExportException;
import net.cua.export.entity.WaterMark;
import net.cua.export.entity.e7.style.*;
import net.cua.export.entity.e7.style.Font;
import net.cua.export.manager.Const;
import net.cua.export.manager.RelManager;
import net.cua.export.processor.IntConversionProcessor;
import net.cua.export.processor.StyleProcessor;
import net.cua.export.util.ExtBufferedWriter;
import net.cua.export.util.FileUtil;
import net.cua.export.util.StringUtil;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import java.io.*;
import java.nio.ByteBuffer;
import java.nio.channels.FileChannel;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Timestamp;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Date;

import static net.cua.export.util.DateUtil.toDateTimeValue;
import static net.cua.export.util.DateUtil.toDateValue;

/**
 * Created by guanquan.wang on 2017/9/26.
 */
@TopNS(prefix = {"", "r"}, value = "worksheet", uri = {Const.SCHEMA_MAIN, Const.Relationship.RELATIONSHIP})
public abstract class Sheet {
    Logger logger = LogManager.getLogger(getClass());
    protected Workbook workbook;

    protected String name;
    protected HeadColumn[] headColumns;
    protected WaterMark waterMark;
    protected RelManager relManager;
    protected int id;
    private int autoSize;
    private double width = 20;
    private int headInfoLen, baseInfoLen;
    protected int rows;
    private boolean hidden;

    private int headStyle;

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
    public Sheet(Workbook workbook, String name, final HeadColumn[] headColumns) {
        this.workbook = workbook;
        this.name = name;
        this.headColumns = headColumns;
        for (int i = 0; i < headColumns.length; i++) {
            headColumns[i].sst = workbook.getStyles();
        }
        relManager = new RelManager();
    }

    public Sheet(Workbook workbook, String name, WaterMark waterMark, final HeadColumn[] headColumns) {
        this.workbook = workbook;
        this.name = name;
        this.headColumns = headColumns;
        for (int i = 0; i < headColumns.length; i++) {
            headColumns[i].sst = workbook.getStyles();
        }
        this.waterMark = waterMark;
        relManager = new RelManager();
    }

    public static class HeadColumn {
        public static final int TYPE_NORMAL = 0
                , TYPE_PARENTAGE = 1 // 百分比
                , TYPE_RMB = 2; // 人民币
        private String key; // Map的主键,object的属性名
        private String name; // 列名
        private Class clazz;     // 列类型
        private boolean share; // 字符串是否共享
        private int type; // 0: 正常显示 1:显示百分比 2:显示人民币
        private IntConversionProcessor processor;
        private StyleProcessor styleProcessor;
        private int cellStyle = -1; // 未设定
        private double width;
        private Object o;
        private Styles sst;

        public void setO(Object o) {
            this.o = o;
        }

        public String getKey() {
            return key;
        }

        public void setKey(String key) {
            this.key = key;
        }

        public IntConversionProcessor getProcessor() {
            return processor;
        }

        public StyleProcessor getStyleProcessor() {
            return styleProcessor;
        }

        public Object getO() {
            return o;
        }

        protected void setSst(Styles styles) {
            this.sst = styles;
        }
        public HeadColumn() {}
        public HeadColumn(String name, Class clazz) {
            this(name, clazz, false);
        }

        public HeadColumn(String name, String key) {
            this(name, key, false);
        }
        public HeadColumn(String name, String key, Class<?> clazz) {
            this(name, key, false);
            this.clazz = clazz;
        }
        public HeadColumn(String name, Class clazz, IntConversionProcessor processor) {
            this(name, clazz, processor, false);
        }
        public HeadColumn(String name, String key, IntConversionProcessor processor) {
            this(name, key, processor, false);
        }

        public HeadColumn(String name, Class clazz, boolean share) {
            this.name = name;
            this.clazz = clazz;
            this.share = share;
        }

        public HeadColumn(String name, String key, boolean share) {
            this.name = name;
            this.key = key;
            this.share = share;
        }

        public HeadColumn(String name, Class clazz, IntConversionProcessor processor, boolean share) {
            this(name, clazz, share);
            this.processor = processor;
        }

        public HeadColumn(String name, String key, IntConversionProcessor processor, boolean share) {
            this(name, key, share);
            this.processor = processor;
        }

        public HeadColumn(String name, Class clazz, int cellStyle) {
            this(name, clazz, cellStyle, false);
        }

        public HeadColumn(String name, String key, int cellStyle) {
            this(name, key, cellStyle, false);
        }

        public HeadColumn(String name, Class clazz, int cellStyle, boolean share) {
            this(name, clazz, share);
            this.cellStyle = cellStyle;
        }

        public HeadColumn(String name, String key, int cellStyle, boolean share) {
            this(name, key, share);
            this.cellStyle = cellStyle;
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

        public HeadColumn setProcessor(IntConversionProcessor processor) {
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

        int getDefaultHorizontal() {
            int horizontal;
            if (clazz == int.class || clazz == Integer.class
                    || clazz == short.class || clazz == Short.class
                    || clazz == byte.class || clazz == Byte.class
                    || clazz == long.class || clazz == Long.class
                    ) {
                horizontal = Horizontals.RIGHT;
            } else if (clazz == Date.class || clazz == Timestamp.class
                    || clazz == LocalDate.class || clazz == LocalDateTime.class
                    || clazz == char.class || clazz == Character.class) {
                horizontal = Horizontals.CENTER;
            } else {
                horizontal = Horizontals.LEFT;
            }
            return horizontal;
        }

        public HeadColumn setCellStyle(Font font) {
            this.cellStyle =  sst.of(
                    (font != null ? sst.addFont(font) : 0)
                            | Verticals.CENTER
                            | getDefaultHorizontal());
            return this;
        }

        public HeadColumn setCellStyle(Font font, int horizontal) {
            this.cellStyle =  sst.of(
                    (font != null ? sst.addFont(font) : 0)
                            | Verticals.CENTER
                            | horizontal);
            return this;
        }

        public HeadColumn setCellStyle(Font font, Border border) {
            this.cellStyle =  sst.of(
                    (font != null ? sst.addFont(font) : 0)
                            | (border != null ? sst.addBorder(border) : 0)
                            | Verticals.CENTER
                            | getDefaultHorizontal());
            return this;
        }

        public HeadColumn setCellStyle(Font font, Border border, int horizontal) {
            this.cellStyle =  sst.of(
                    (font != null ? sst.addFont(font) : 0)
                            | (border != null ? sst.addBorder(border) : 0)
                            | Verticals.CENTER
                            | horizontal);
            return this;
        }

        public HeadColumn setCellStyle(Font font, Fill fill, Border border) {
            this.cellStyle =  sst.of(
                    (font != null ? sst.addFont(font) : 0)
                            | (fill != null ? sst.addFill(fill) : 0)
                            | (border != null ? sst.addBorder(border) : 0)
                            | Verticals.CENTER
                            | getDefaultHorizontal());
            return this;
        }

        public HeadColumn setCellStyle(Font font, Fill fill, Border border, int horizontal) {
            this.cellStyle =  sst.of(
                    (font != null ? sst.addFont(font) : 0)
                            | (fill != null ? sst.addFill(fill) : 0)
                            | (border != null ? sst.addBorder(border) : 0)
                            | Verticals.CENTER
                            | horizontal);
            return this;
        }

        public HeadColumn setCellStyle(Font font, Fill fill, Border border, int vertical, int horizontal) {
            this.cellStyle =  sst.of(
                            (font != null ? sst.addFont(font) : 0)
                            | (fill != null ? sst.addFill(fill) : 0)
                            | (border != null ? sst.addBorder(border) : 0)
                            | vertical
                            | horizontal);
            return this;
        }

        public HeadColumn setCellStyle(NumFmt numFmt, Font font, Fill fill, Border border, int vertical, int horizontal) {
            this.cellStyle =  sst.of(
                    (numFmt != null ? sst.addNumFmt(numFmt) : 0)
                            | (font != null ? sst.addFont(font) : 0)
                            | (fill != null ? sst.addFill(fill) : 0)
                            | (border != null ? sst.addBorder(border) : 0)
                            | vertical
                            | horizontal);
            return this;
        }

        public HeadColumn setShare(boolean share) {
            this.share = share;
            return this;
        }

        protected int getCellStyle(Class clazz) {
            int style;
            if (clazz == String.class) {
                style = Styles.defaultStringBorderStyle();
            } else if (clazz == java.sql.Date.class
                    || clazz == java.util.Date.class) {
                style = Styles.defaultDateBorderStyle();
            } else if (clazz == java.sql.Timestamp.class) {
                style = Styles.defaultTimestampBorderStyle();
            } else if (clazz == int.class || clazz == Integer.class
                    || clazz == long.class || clazz == Long.class
                    || clazz == byte.class || clazz == Byte.class
                    || clazz == short.class || clazz == Short.class
                    ) {
                style = Styles.defaultIntBorderStyle();
                switch (type) {
                    case TYPE_PARENTAGE: // 百分比显示
                        style = Styles.clearNumfmt(style) | sst.addNumFmt(new NumFmt("0%_);[Red]\\(0%\\)"));
                        break;
                    case TYPE_RMB: // 显示人民币
                        style = Styles.clearNumfmt(style) | sst.addNumFmt(new NumFmt("¥0_);[Red]\\(¥0\\)"));
                        break;
                    case TYPE_NORMAL: // 正常显示数字
                        break;
                    default:
                }
            } else if (clazz == double.class || clazz == Double.class
                    || clazz == float.class || clazz == Float.class
                    ) {
                style = Styles.defaultDoubleBorderStyle();
                switch (type) {
                    case TYPE_PARENTAGE: // 百分比显示
                        style= Styles.clearNumfmt(style) | sst.addNumFmt(new NumFmt("0.00%_);[Red]\\(0.00%\\)"));
                        break;
                    case TYPE_RMB: // 显示人民币
                        style = Styles.clearNumfmt(style) | sst.addNumFmt(new NumFmt("¥0.00_);[Red]\\(¥0.00\\)"));
                        break;
                    case TYPE_NORMAL: // 正常显示数字
                        break;
                default:
                }
            } else if (clazz == boolean.class || clazz == Boolean.class
                    || clazz == char.class || clazz == Character.class
                    ) {
                style = Styles.clearHorizontal(Styles.defaultStringBorderStyle()) | Horizontals.CENTER;
            } else {
                style = 0; // Auto-style
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
        for (int i = 0; i < headColumns.length; i++) {
            headColumns[i].sst = workbook.getStyles();
        }
    }

    public WaterMark getWaterMark() {
        return waterMark;
    }

    public void setWaterMark(WaterMark waterMark) {
        this.waterMark = waterMark;
    }

    public boolean isHidden() {
        return hidden;
    }
    public Sheet hidden() {
        // TODO sheet hidden
        this.hidden = true;
        return this;
    }
    /**
     * abstract method close
     */
    public abstract void close();

    public void addRel(Relationship rel) {
        relManager.add(rel);
    }

    public abstract void writeTo(Path xl) throws IOException, ExportException;

    protected String getFileName() {
        return "sheet" + id + Const.Suffix.XML;
    }

    /**
     * 写worksheet头部
     * @param bw
     */
    protected void writeBefore(ExtBufferedWriter bw) throws IOException {
        StringBuilder buf = new StringBuilder(Const.EXCEL_XML_DECLARATION);
        // Declaration
        buf.append(Const.lineSeparator); // new line
        // Root node
        if (getClass().isAnnotationPresent(TopNS.class)) {
            TopNS topNS = getClass().getAnnotation(TopNS.class);
            buf.append('<').append(topNS.value());
            String[] prefixs = topNS.prefix(), urls = topNS.uri();
            for (int i = 0, len = prefixs.length; i < len; ) {
                buf.append(" xmlns");
                if (prefixs[i] != null && !prefixs[i].isEmpty()) {
                    buf.append(':').append(prefixs[i]);
                }
                buf.append("=\"").append(urls[i]);
                if (++i < len) {
                    buf.append('"');
                }
            }
        } else {
            buf.append("<worksheet xmlns=\"").append(Const.SCHEMA_MAIN);
        }
        buf.append("\">");

        // Dimension
        buf.append("<dimension ref=\"A1\"/>");
        headInfoLen = buf.length() - 3;

        // SheetViews default value
        buf.append("<sheetViews><sheetView workbookViewId=\"0\"");
        if (id == 1) { // Default select the first worksheet
            buf.append(" tabSelected=\"1\"");
        }
        buf.append("/></sheetViews>");

        // Default format
        buf.append("<sheetFormatPr defaultRowHeight=\"13.5\" baseColWidth=\"");
        buf.append((int)width);
        buf.append("\"/>");

        baseInfoLen = buf.length() - headInfoLen;
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

    public Sheet setHeadStyle(Font font, Fill fill, Border border) {
        return setHeadStyle(null, font, fill, border, Verticals.CENTER, Horizontals.CENTER);
    }

    public Sheet setHeadStyle(Font font, Fill fill, Border border, int vertical, int horizontal) {
        return setHeadStyle(null, font, fill, border, vertical, horizontal);
    }

    public Sheet setHeadStyle(NumFmt numFmt, Font font, Fill fill, Border border, int vertical, int horizontal) {
        Styles styles = workbook.getStyles();
        headStyle = styles.of(
                (numFmt != null ? styles.addNumFmt(numFmt) : 0)
                        | (font != null ? styles.addFont(font) : 0)
                        | (fill != null ? styles.addFill(fill) : 0)
                        | (border != null ? styles.addBorder(border) : 0)
                        | vertical
                        | horizontal);
        return this;
    }

    public Sheet setHeadStyle(int style) {
        headStyle = style;
        return this;
    }

    private int defaultHeadStyle() {
        if (headStyle == 0) {
            Styles styles = workbook.getStyles();
            headStyle = styles.of(styles.addFont(Font.parse("bold 11 微软雅黑 white"))
                            | styles.addFill(Fill.parse("solid #666699"))
                            | styles.addBorder(Border.parse("thin black"))
                            | Verticals.CENTER
                            | Horizontals.CENTER);
        }
        return headStyle;
    }

    /**
     * 写尾部
     * @param bw
     */
    protected void writeAfter(ExtBufferedWriter bw) throws IOException {
        // End target --sheetData
        bw.write("</sheetData>");

        // background image
        if (waterMark != null) {
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
    protected void writeRow(ResultSet rs, ExtBufferedWriter bw) throws IOException, SQLException {
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

            // t n=numeric (default), s=string, b=boolean, str=function string
            // TODO function <f ca="1" or t="shared" ref="O10:O15" si="0" ... si="10"></f>
            if (hc.clazz == String.class) {
                String s = rs.getString(i + 1);
                writeString(bw, s, i);
            }
            else if (hc.clazz == java.util.Date.class
                    || hc.clazz == java.sql.Date.class) {
                java.sql.Date date = rs.getDate(i + 1);
                writeDate(bw, date, i);
            }
            else if (hc.clazz == java.sql.Timestamp.class) {
                Timestamp ts = rs.getTimestamp(i + 1);
                writeTimestamp(bw, ts, i);
            }
            else if (hc.clazz == int.class || hc.clazz == Integer.class
                    || hc.clazz == char.class || hc.clazz == Character.class
                    || hc.clazz == byte.class || hc.clazz == Byte.class
                    || hc.clazz == short.class || hc.clazz == Short.class
                    ) {
                int n = rs.getInt(i + 1);
                writeInt(bw, n, i);
            }
            else if (hc.clazz == long.class || hc.clazz == Long.class) {
                long l = rs.getLong(i + 1);
                writeLong(bw, l, i);
            }
            else if (hc.clazz == double.class || hc.clazz == Double.class
                    || hc.clazz == float.class || hc.clazz == Float.class
                    ) {
                double d = rs.getDouble(i + 1);
                writeDouble(bw, d, i);
            } else if (hc.clazz == boolean.class || hc.clazz == Boolean.class) {
                boolean bool = rs.getBoolean(i + 1);
                writeBoolean(bw, bool, i);
            } else {
                Object o = rs.getObject(i + 1);
                if (o != null) {
                    writeString(bw, o.toString(), i);
                } else {
                    writeNull(bw, i);
                }
            }
        }
        bw.write("</row>");
    }

    /**
     * 写行数据
     * @param rs ResultSet
     * @param bw
     */
    protected void writeRowAutoSize(ResultSet rs, ExtBufferedWriter bw) throws IOException, SQLException {
        int r = ++rows;
        final int len = headColumns.length;
        bw.write("<row r=\"");
        bw.writeInt(r);
        bw.write("\" ht=\"16.5\" spans=\"1:");
        bw.writeInt(len);
        bw.write("\">");

        for (int i = 0; i < len; i++) {
            HeadColumn hc = headColumns[i];
            // t n=numeric (default), s=string, b=boolean, str=function string
            // TODO function <f ca="1" or t="shared" ref="O10:O15" si="0" ... si="10"></f>
            if (hc.clazz == String.class) {
                String s = rs.getString(i + 1);
                writeStringAutoSize(bw, s, i);
            }
            else if (hc.clazz == java.util.Date.class
                    || hc.clazz == java.sql.Date.class) {
                java.sql.Date date = rs.getDate(i + 1);
                writeDate(bw, date, i);
            }
            else if (hc.clazz == java.sql.Timestamp.class) {
                Timestamp ts = rs.getTimestamp(i + 1);
                writeTimestamp(bw, ts, i);
            }
            else if (hc.clazz == int.class || hc.clazz == Integer.class
                    || hc.clazz == char.class || hc.clazz == Character.class
                    || hc.clazz == byte.class || hc.clazz == Byte.class
                    || hc.clazz == short.class || hc.clazz == Short.class
                    ) {
                int n = rs.getInt(i + 1);
                writeIntAutoSize(bw, n, i);
            }
            else if (hc.clazz == long.class || hc.clazz == Long.class) {
                long l = rs.getLong(i + 1);
                writeLong(bw, l, i);
            }
            else if (hc.clazz == double.class || hc.clazz == Double.class
                    || hc.clazz == float.class || hc.clazz == Float.class
                    ) {
                double d = rs.getDouble(i + 1);
                writeDouble(bw, d, i);
            } else if (hc.clazz == boolean.class || hc.clazz == Boolean.class) {
                boolean bool = rs.getBoolean(i + 1);
                writeBoolean(bw, bool, i);
            } else {
                Object o = rs.getObject(i + 1);
                if (o != null) {
                    writeStringAutoSize(bw, o.toString(), i);
                } else {
                    writeNull(bw, i);
                }
            }
        }
        bw.write("</row>");
    }

    protected int getStyleIndex(HeadColumn hc, Object o) {
        int style = hc.getCellStyle(), styleIndex = hc.sst.of(style);
        if (hc.styleProcessor != null) {
            style = hc.styleProcessor.build(o, style, hc.sst);
            styleIndex = hc.sst.of(style);
        }
        return styleIndex;
    }

    protected void writeString(ExtBufferedWriter bw, String s, int column) throws IOException {
        writeString(bw, s, column, s);
    }

    private void writeString(ExtBufferedWriter bw, String s, int column, Object o) throws IOException {
        HeadColumn hc = headColumns[column];
        int styleIndex = getStyleIndex(hc, o);
        bw.write("<c r=\"");
        bw.write(int2Col(column + 1));
        bw.writeInt(rows);
        int i;
        if (StringUtil.isEmpty(s)) {
            bw.write("\" s=\"");
            bw.writeInt(styleIndex);
            bw.write("\"/>");
        }
        else if (hc.isShare() && (i = workbook.getSst().get(s)) >= 0) {
            bw.write("\" t=\"s\" s=\"");
            bw.writeInt(styleIndex);
            bw.write("\"><v>");
            bw.writeInt(i);
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

    protected void writeStringAutoSize(ExtBufferedWriter bw, String s, int column) throws IOException {
        writeStringAutoSize(bw, s, column, s);
    }

    protected void writeStringAutoSize(ExtBufferedWriter bw, String s, int column, Object o) throws IOException {
        HeadColumn hc = headColumns[column];
        int styleIndex = getStyleIndex(hc, o);
        bw.write("<c r=\"");
        bw.write(int2Col(column + 1));
        bw.writeInt(rows);
        if (StringUtil.isEmpty(s)) {
            bw.write("\" s=\"");
            bw.writeInt(styleIndex);
            bw.write("\"/>");
        } else {
            int i;
            if (hc.isShare() && (i = workbook.getSst().get(s)) >= 0) {
                bw.write("\" t=\"s\" s=\"");
                bw.writeInt(styleIndex);
                bw.write("\"><v>");
                bw.writeInt(i);
                bw.write("</v></c>");
            } else {
                bw.write("\" t=\"inlineStr\" s=\"");
                bw.writeInt(styleIndex);
                bw.write("\"><is><t>");
                bw.write(s);
                bw.write("</t></is></c>");
            }
            int ln = s.getBytes("GB2312").length; // TODO 计算
            if (hc.width == 0 && (hc.o == null || (int) hc.o < ln)) {
                hc.o = ln;
            }
        }
    }

    protected void writeDate(ExtBufferedWriter bw, Date date, int column) throws IOException {
        writeDate(bw, date, column, date);
    }

    protected void writeDate(ExtBufferedWriter bw, Date date, int column, Object o) throws IOException {
        int styleIndex = getStyleIndex(headColumns[column], o);
        bw.write("<c r=\"");
        bw.write(int2Col(column + 1));
        bw.writeInt(rows);
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

    protected void writeTimestamp(ExtBufferedWriter bw, Timestamp ts, int column) throws IOException {
        writeTimestamp(bw, ts, column, ts);
    }

    protected void writeTimestamp(ExtBufferedWriter bw, Timestamp ts, int column, Object o) throws IOException {
        int styleIndex = getStyleIndex(headColumns[column], o);
        bw.write("<c r=\"");
        bw.write(int2Col(column + 1));
        bw.writeInt(rows);
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

    protected void writeInt(ExtBufferedWriter bw, int n, int column) throws IOException {
        HeadColumn hc = headColumns[column];
        if (hc.processor == null) {
            writeInt0(bw, n, column);
        } else {
            Object o = hc.processor.conversion(n);
            if (o != null) {
                Class<?> clazz = o.getClass();
                boolean blockOrDefault = hc.cellStyle == -1 || hc.cellStyle == Styles.defaultIntBorderStyle();
                if (clazz == String.class) {
                    if (blockOrDefault) {
                        hc.cellStyle = hc.getCellStyle(String.class);
                    }
                    writeString(bw, o.toString(), column, n);
                }
                else if (clazz == int.class || clazz == Integer.class
                        || clazz == byte.class || clazz == Byte.class
                        || clazz == short.class || clazz == Short.class
                        ) {
                    n = ((Integer) o).intValue();
                    writeInt0(bw, n, column, n);
                }
                else if (clazz == char.class || clazz == Character.class) {
                    if (blockOrDefault) {
                        hc.cellStyle = Styles.defaultCharBorderStyle();
                    }
                    char c = ((Character) o).charValue();
                    writeChar0(bw, c, column, n);
                }
                else if (clazz == long.class || clazz == Long.class) {
                    long l = ((Long) o).longValue();
                    writeLong(bw, l, column, n);
                }
                else if (clazz == java.util.Date.class
                        || clazz == java.sql.Date.class) {
                    if (blockOrDefault) {
                        hc.cellStyle = hc.getCellStyle(java.util.Date.class);
                    }
                    writeDate(bw, (Date) o, column, n);
                }
                else if (clazz == java.sql.Timestamp.class) {
                    if (blockOrDefault) {
                        hc.cellStyle = hc.getCellStyle(java.sql.Timestamp.class);
                    }
                    writeTimestamp(bw, (Timestamp) o, column, n);
                }
                else if (hc.clazz == double.class || hc.clazz == Double.class
                        || hc.clazz == float.class || hc.clazz == Float.class
                        ) {
                    if (blockOrDefault) {
                        hc.cellStyle = hc.getCellStyle(double.class);
                    }
                    writeDouble(bw, ((Double) o).doubleValue(), column, n);
                }
                else if (hc.clazz == boolean.class || hc.clazz == Boolean.class) {
                    if (blockOrDefault) {
                        hc.cellStyle = hc.getCellStyle(boolean.class);
                    }
                    boolean bool = ((Boolean) o).booleanValue();
                    writeBoolean(bw, bool, column, n);
                }
                else {
                    if (blockOrDefault) {
                        hc.cellStyle = hc.getCellStyle(String.class);
                    }
                    writeString(bw, o.toString(), column, n);
                }
            }
            else {
                writeNull(bw, column);
            }
        }
    }

    protected void writeIntAutoSize(ExtBufferedWriter bw, int n, int column) throws IOException {
        HeadColumn hc = headColumns[column];
        if (hc.processor == null) {
            writeInt0(bw, n, column);
        } else {
            Object o = hc.processor.conversion(n);
            if (o != null) {
                Class<?> clazz = o.getClass();
                boolean blockOrDefault = hc.cellStyle == -1 || hc.cellStyle == Styles.defaultIntBorderStyle();
                if (clazz == String.class) {
                    if (blockOrDefault) {
                        hc.cellStyle = hc.getCellStyle(String.class);
                    }
                    writeStringAutoSize(bw, o.toString(), column, n);
                }
                else if (clazz == int.class || clazz == Integer.class
                        || clazz == char.class || clazz == Character.class
                        || clazz == byte.class || clazz == Byte.class
                        || clazz == short.class || clazz == Short.class
                        ) {
                    int nn = ((Integer) o).intValue();
                    writeInt0(bw, nn, column, n);
                }
                else if (clazz == char.class || clazz == Character.class) {
                    if (blockOrDefault) {
                        hc.cellStyle = Styles.defaultCharBorderStyle();
                    }
                    char c = ((Character) o).charValue();
                    writeChar0(bw, c, column, n);
                }
                else if (clazz == long.class || clazz == Long.class) {
                    long l = ((Long) o).longValue();
                    writeLong(bw, l, column, n);
                }
                else if (clazz == java.util.Date.class
                        || clazz == java.sql.Date.class) {
                    if (blockOrDefault) {
                        hc.cellStyle = hc.getCellStyle(java.util.Date.class);
                    }
                    writeDate(bw, (Date) o, column, n);
                }
                else if (clazz == java.sql.Timestamp.class) {
                    if (blockOrDefault) {
                        hc.cellStyle = hc.getCellStyle(java.sql.Timestamp.class);
                    }
                    writeTimestamp(bw, (Timestamp) o, column, n);
                }
                else if (hc.clazz == double.class || hc.clazz == Double.class
                        || hc.clazz == float.class || hc.clazz == Float.class
                        ) {
                    if (blockOrDefault) {
                        hc.cellStyle = hc.getCellStyle(double.class);
                    }
                    writeDouble(bw, ((Double) o).doubleValue(), column, n);
                }
                else if (hc.clazz == boolean.class || hc.clazz == Boolean.class) {
                    if (blockOrDefault) {
                        hc.cellStyle = hc.getCellStyle(boolean.class);
                    }
                    boolean bool = ((Boolean) o).booleanValue();
                    writeBoolean(bw, bool, column, n);
                }
                else {
                    if (blockOrDefault) {
                        hc.cellStyle = hc.getCellStyle(String.class);
                    }
                    writeStringAutoSize(bw, o.toString(), column, n);
                }
            }
            else {
                writeNull(bw, column);
            }
        }
    }

    protected void writeChar(ExtBufferedWriter bw, char c, int column) throws IOException {
        HeadColumn hc = headColumns[column];
        if (hc.processor == null) {
            writeChar0(bw, c, column);
        } else {
            Object o = hc.processor.conversion(c);
            if (o != null) {
                Class<?> clazz = o.getClass();
                boolean blockOrDefault = hc.cellStyle == -1 || hc.cellStyle == Styles.defaultCharBorderStyle();
                if (clazz == String.class) {
                    if (blockOrDefault) {
                        hc.cellStyle = hc.getCellStyle(String.class);
                    }
                    writeString(bw, o.toString(), column, c);
                }
                else if (clazz == int.class || clazz == Integer.class
                        || clazz == byte.class || clazz == Byte.class
                        || clazz == short.class || clazz == Short.class
                        ) {
                    if (blockOrDefault) {
                        hc.cellStyle = hc.getCellStyle(int.class);
                    }
                    int n = ((Integer) o).intValue();
                    writeInt0(bw, n, column, c);
                }
                else if (clazz == char.class || clazz == Character.class) {
                    char cc = ((Character) o).charValue();
                    writeChar0(bw, cc, column, c);
                }
                else if (clazz == long.class || clazz == Long.class) {
                    if (blockOrDefault) {
                        hc.cellStyle = hc.getCellStyle(long.class);
                    }
                    long l = ((Long) o).longValue();
                    writeLong(bw, l, column, c);
                }
                else if (clazz == java.util.Date.class
                        || clazz == java.sql.Date.class) {
                    if (blockOrDefault) {
                        hc.cellStyle = hc.getCellStyle(java.util.Date.class);
                    }
                    writeDate(bw, (Date) o, column, c);
                }
                else if (clazz == java.sql.Timestamp.class) {
                    if (blockOrDefault) {
                        hc.cellStyle = hc.getCellStyle(java.sql.Timestamp.class);
                    }
                    writeTimestamp(bw, (Timestamp) o, column, c);
                }
                else if (hc.clazz == double.class || hc.clazz == Double.class
                        || hc.clazz == float.class || hc.clazz == Float.class
                        ) {
                    if (blockOrDefault) {
                        hc.cellStyle = hc.getCellStyle(double.class);
                    }
                    writeDouble(bw, ((Double) o).doubleValue(), column, c);
                }
                else if (hc.clazz == boolean.class || hc.clazz == Boolean.class) {
                    if (blockOrDefault) {
                        hc.cellStyle = hc.getCellStyle(boolean.class);
                    }
                    boolean bool = ((Boolean) o).booleanValue();
                    writeBoolean(bw, bool, column, c);
                }
                else {
                    if (blockOrDefault) {
                        hc.cellStyle = hc.getCellStyle(String.class);
                    }
                    writeString(bw, o.toString(), column, c);
                }
            }
            else {
                writeNull(bw, column);
            }
        }
    }

    protected void writeCharAutoSize(ExtBufferedWriter bw, char c, int column) throws IOException {
        HeadColumn hc = headColumns[column];
        if (hc.processor == null) {
            writeChar0(bw, c, column);
        } else {
            Object o = hc.processor.conversion(c);
            if (o != null) {
                Class<?> clazz = o.getClass();
                boolean blockOrDefault = hc.cellStyle == -1 || hc.cellStyle == Styles.defaultCharBorderStyle();
                if (clazz == String.class) {
                    if (blockOrDefault) {
                        hc.cellStyle = hc.getCellStyle(String.class);
                    }
                    writeStringAutoSize(bw, o.toString(), column);
                }
                else if (clazz == int.class || clazz == Integer.class
                        || clazz == byte.class || clazz == Byte.class
                        || clazz == short.class || clazz == Short.class
                        ) {
                    if (blockOrDefault) {
                        hc.cellStyle = hc.getCellStyle(int.class);
                    }
                    int n = ((Integer) o).intValue();
                    writeInt0(bw, n, column, c);
                }
                else if (clazz == char.class || clazz == Character.class) {
                    char cc = ((Character) o).charValue();
                    writeChar0(bw, cc, column, c);
                }
                else if (clazz == long.class || clazz == Long.class) {
                    if (blockOrDefault) {
                        hc.cellStyle = hc.getCellStyle(long.class);
                    }
                    long l = ((Long) o).longValue();
                    writeLong(bw, l, column);
                }
                else if (clazz == java.util.Date.class
                        || clazz == java.sql.Date.class) {
                    if (blockOrDefault) {
                        hc.cellStyle = hc.getCellStyle(java.util.Date.class);
                    }
                    writeDate(bw, (Date) o, column);
                }
                else if (clazz == java.sql.Timestamp.class) {
                    if (blockOrDefault) {
                        hc.cellStyle = hc.getCellStyle(java.sql.Timestamp.class);
                    }
                    writeTimestamp(bw, (Timestamp) o, column);
                }
                else if (hc.clazz == double.class || hc.clazz == Double.class
                        || hc.clazz == float.class || hc.clazz == Float.class
                        ) {
                    if (blockOrDefault) {
                        hc.cellStyle = hc.getCellStyle(double.class);
                    }
                    writeDouble(bw, ((Double) o).doubleValue(), column);
                }
                else if (hc.clazz == boolean.class || hc.clazz == Boolean.class) {
                    if (blockOrDefault) {
                        hc.cellStyle = hc.getCellStyle(boolean.class);
                    }
                    boolean bool = ((Boolean) o).booleanValue();
                    writeBoolean(bw, bool, column);
                }
                else {
                    if (blockOrDefault) {
                        hc.cellStyle = hc.getCellStyle(String.class);
                    }
                    writeStringAutoSize(bw, o.toString(), column);
                }
            } else {
                writeNull(bw, column);
            }
        }
    }
    private void writeInt0(ExtBufferedWriter bw, int n, int column) throws IOException {
        writeInt0(bw, n, column, n);
    }

    private void writeInt0(ExtBufferedWriter bw, int n, int column, Object o) throws IOException {
        int styleIndex = getStyleIndex(headColumns[column], o);
        bw.write("<c r=\"");
        bw.write(int2Col(column + 1));
        bw.writeInt(rows);
        bw.write("\" s=\"");
        bw.writeInt(styleIndex);
        bw.write("\"><v>");
        bw.writeInt(n);
        bw.write("</v></c>");
    }

    private void writeChar0(ExtBufferedWriter bw, char c, int column) throws IOException {
        writeChar0(bw, c, column, c);
    }

    private void writeChar0(ExtBufferedWriter bw, char c, int column, Object o) throws IOException {
        int styleIndex = getStyleIndex(headColumns[column], o);
        bw.write("<c r=\"");
        bw.write(int2Col(column + 1));
        bw.writeInt(rows);
        bw.write("\" t=\"s\" s=\"");
        bw.writeInt(styleIndex);
        bw.write("\"><v>");
        bw.writeInt(workbook.getSst().get(c));
        bw.write("</v></c>");
    }

    protected void writeLong(ExtBufferedWriter bw, long l, int column) throws IOException {
        writeLong(bw, l, column, l);
    }

    protected void writeLong(ExtBufferedWriter bw, long l, int column, Object o) throws IOException {
        int styleIndex = getStyleIndex(headColumns[column], o);
        bw.write("<c r=\"");
        bw.write(int2Col(column + 1));
        bw.writeInt(rows);
        bw.write("\" s=\"");
        bw.writeInt(styleIndex);
        bw.write("\"><v>");
        bw.write(l);
        bw.write("</v></c>");
    }

    protected void writeDouble(ExtBufferedWriter bw, double d, int column) throws IOException {
        writeDouble(bw, d, column, d);
    }

    protected void writeDouble(ExtBufferedWriter bw, double d, int column, Object o) throws IOException {
        int styleIndex = getStyleIndex(headColumns[column], o);
        bw.write("<c r=\"");
        bw.write(int2Col(column + 1));
        bw.writeInt(rows);
        bw.write("\" s=\"");
        bw.writeInt(styleIndex);
        bw.write("\"><v>");
        bw.write(d);
        bw.write("</v></c>");
    }

    protected void writeBoolean(ExtBufferedWriter bw, boolean bool, int column) throws IOException {
        writeBoolean(bw, bool, column, bool);
    }

    protected void writeBoolean(ExtBufferedWriter bw, boolean bool, int column, Object o) throws IOException {
        int styleIndex = getStyleIndex(headColumns[column], o);
        bw.write("<c r=\"");
        bw.write(int2Col(column + 1));
        bw.writeInt(rows);
        bw.write("\" t=\"b\" s=\"");
        bw.writeInt(styleIndex);
        bw.write("\"><v>");
        bw.writeInt(bool ? 1 : 0);
        bw.write("</v></c>");
    }

    protected void writeNull(ExtBufferedWriter bw, int column) throws IOException {
        int styleIndex = getStyleIndex(headColumns[column], null);
        bw.write("<c r=\"");
        bw.write(int2Col(column + 1));
        bw.writeInt(rows);
        bw.write("\" s=\"");
        bw.writeInt(styleIndex);
        bw.write("\"/>");
    }

    /**
     * 写空行数据
     * @param bw
     */
    protected void writeEmptyRow(ExtBufferedWriter bw) throws IOException {
        // Row number
        int r = ++rows;
        final int len = headColumns.length;
        bw.write("<row r=\"");
        bw.writeInt(r);
        bw.write("\" ht=\"16.5\" spans=\"1:");
        bw.writeInt(len);
        bw.write("\">");

        Styles styles = workbook.getStyles();
        for (int i = 1; i <= len; i++) {
            HeadColumn hc = headColumns[i - 1];
            bw.write("<c r=\"");
            bw.write(int2Col(i));
            bw.writeInt(r);

            int style = hc.getCellStyle();
            int styleIndex = styles.of(style);
            bw.write("\" s=\"");
            bw.writeInt(styleIndex);
            bw.write("\"/>");

            if (hc.getO() == null) {
                hc.setO(hc.getName().getBytes("GB2312").length);
            }
        }
        bw.write("</row>");
    }

    protected  void autoColumnSize(File sheet) throws IOException {
        // resize each column width ...
        File temp = new File(sheet.getParent(), sheet.getName() + ".temp");
        if (!sheet.renameTo(temp)) {
            Files.move(sheet.toPath(), temp.toPath(), StandardCopyOption.REPLACE_EXISTING);
        }

        try (FileChannel inChannel = new FileInputStream(temp).getChannel();
             FileChannel outChannel = new FileOutputStream(sheet).getChannel()) {

            inChannel.transferTo(0, headInfoLen, outChannel);
            ByteBuffer buffer = ByteBuffer.allocate(baseInfoLen);
            inChannel.read(buffer, headInfoLen);
            buffer.compact();
            byte b;
            if ((b = buffer.get()) == '"') {
                char[] chars = int2Col(headColumns.length);
                String s = ':' + new String(chars) + rows;
                outChannel.write(ByteBuffer.wrap(s.getBytes(Const.UTF_8)));
            }
            buffer.flip();
            buffer.put(b);
            buffer.compact();
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
                        if (hc.o == null) {
                            len = 0;
                        } else {
                            len = (int) hc.o;
                        }
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
                        if (hc.o == null) {
                            len = 0;
                        } else {
                            len = hc.o.toString().getBytes("GB2312").length;
                        }
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

            outChannel.write(ByteBuffer.wrap(buf.toString().getBytes(Const.UTF_8)));
            int start = headInfoLen + baseInfoLen;
            inChannel.transferTo(start, inChannel.size() - start, outChannel);

        } catch (IOException e) {
            throw e;
        } finally {
            boolean delete = temp.delete();
            if (!delete) {
                logger.error("Delete temp file failed.");
            }
        }
    }

    private ThreadLocal<char[][]> cache = ThreadLocal.withInitial(() -> new char[][] {{65}, {65, 65}, {65, 65, 65}});
    protected char[] int2Col(int n) {
        char[][] cache_col = cache.get();
        char[] c; char A = 'A';
        if (n <= 26) {
            c = cache_col[0];
            c[0] = (char) (n - 1 + A);
        } else if (n <= 702) {
            int t = n / 26, w = n % 26;
            if (w == 0) {
                t--;
                w = 26;
            }
            c = cache_col[1];
            c[0] = (char) (t - 1 + A);
            c[1] = (char) (w - 1 + A);
        } else {
            int t = n / 26, tt = t / 26, w = n % 26;
            if (w == 0) {
                t--;
                w = 26;
            }
            c = cache_col[2];
            c[0] = (char) (tt - 1 + A);
            c[1] = (char) (t - 27 + A);
            c[2] = (char) (w - 1 + A);
        }
        return c;
    }

}

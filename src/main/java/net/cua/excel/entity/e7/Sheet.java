package net.cua.excel.entity.e7;

import net.cua.excel.entity.e7.style.*;
import net.cua.excel.manager.Const;
import net.cua.excel.manager.RelManager;
import net.cua.excel.annotation.TopNS;
import net.cua.excel.entity.ExportException;
import net.cua.excel.entity.WaterMark;
import net.cua.excel.processor.IntConversionProcessor;
import net.cua.excel.processor.StyleProcessor;
import net.cua.excel.util.ExtBufferedWriter;
import net.cua.excel.util.StringUtil;
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
import java.util.Date;

import static net.cua.excel.util.DateUtil.toDateTimeValue;
import static net.cua.excel.util.DateUtil.toDateValue;

/**
 * Created by guanquan.wang on 2017/9/26.
 */
@TopNS(prefix = {"", "r"}, value = "worksheet", uri = {Const.SCHEMA_MAIN, Const.Relationship.RELATIONSHIP})
public abstract class Sheet {
    Logger logger = LogManager.getLogger(getClass());
    protected Workbook workbook;

    protected String name;
    protected Column[] columns;
    protected WaterMark waterMark;
    protected RelManager relManager;
    protected int id;
    private int autoSize;
    private double width = 20;
    private int headInfoLen, baseInfoLen;
    protected int rows;
    private boolean hidden;

    private int headStyle, oddStyle;

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
    public Sheet(Workbook workbook, String name, final Column[] columns) {
        this.workbook = workbook;
        this.name = name;
        this.columns = columns;
        for (int i = 0; i < columns.length; i++) {
            columns[i].styles = workbook.getStyles();
        }
        relManager = new RelManager();
    }

    public Sheet(Workbook workbook, String name, WaterMark waterMark, final Column[] columns) {
        this.workbook = workbook;
        this.name = name;
        this.columns = columns;
        for (int i = 0; i < columns.length; i++) {
            columns[i].styles = workbook.getStyles();
        }
        this.waterMark = waterMark;
        relManager = new RelManager();
    }

    public static class Column {
        public static final int TYPE_NORMAL = 0
                , TYPE_PARENTAGE = 1 // 百分比
                , TYPE_RMB = 2; // 人民币
        private String key; // Map的主键,object的属性名
        private String name; // 列名
        private Class<?> clazz;     // 列类型
        private boolean share; // 字符串是否共享
        private int type; // 0: 正常显示 1:显示百分比 2:显示人民币
        private IntConversionProcessor processor;
        private StyleProcessor styleProcessor;
        private int cellStyle = -1; // 未设定
        private double width;
        private Object o;
        private Styles styles;

        public Column setO(Object o) {
            this.o = o;
            return this;
        }

        public String getKey() {
            return key;
        }

        public Column setKey(String key) {
            this.key = key;
            return this;
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
            this.styles = styles;
        }
        public Column() {}
        public Column(String name, Class<?> clazz) {
            this(name, clazz, false);
        }

        public Column(String name, String key) {
            this(name, key, false);
        }
        public Column(String name, String key, Class<?> clazz) {
            this(name, key, false);
            this.clazz = clazz;
        }
        public Column(String name, Class<?> clazz, IntConversionProcessor processor) {
            this(name, clazz, processor, false);
        }
        public Column(String name, String key, IntConversionProcessor processor) {
            this(name, key, processor, false);
        }

        public Column(String name, Class<?> clazz, boolean share) {
            this.name = name;
            this.clazz = clazz;
            this.share = share;
        }

        public Column(String name, String key, boolean share) {
            this.name = name;
            this.key = key;
            this.share = share;
        }

        public Column(String name, Class<?> clazz, IntConversionProcessor processor, boolean share) {
            this(name, clazz, share);
            this.processor = processor;
        }

        public Column(String name, String key, IntConversionProcessor processor, boolean share) {
            this(name, key, share);
            this.processor = processor;
        }

        public Column(String name, Class<?> clazz, int cellStyle) {
            this(name, clazz, cellStyle, false);
        }

        public Column(String name, String key, int cellStyle) {
            this(name, key, cellStyle, false);
        }

        public Column(String name, Class<?> clazz, int cellStyle, boolean share) {
            this(name, clazz, share);
            this.cellStyle = cellStyle;
        }

        public Column(String name, String key, int cellStyle, boolean share) {
            this(name, key, share);
            this.cellStyle = cellStyle;
        }

        public Column setWidth(double width) {
            if (width < 0.00000001) {
                throw new RuntimeException("Width " + width + " less than 0.");
            }
            this.width = width;
            return this;
        }

        public boolean isShare() {
            return share;
        }

        public Column setType(int type) {
            this.type = type;
            return this;
        }

        public String getName() {
            return name;
        }

        public Column setName(String name) {
            this.name = name;
            return this;
        }

        public Class<?> getClazz() {
            return clazz;
        }

        public Column setClazz(Class<?> clazz) {
            this.clazz = clazz;
            return this;
        }

        public Column setProcessor(IntConversionProcessor processor) {
            this.processor = processor;
            return this;
        }

        public Column setStyleProcessor(StyleProcessor styleProcessor) {
            this.styleProcessor = styleProcessor;
            return this;
        }

        public double getWidth() {
            return width;
        }

        public Column setCellStyle(int cellStyle) {
            // TODO when style not exists
            this.cellStyle = cellStyle;
            return this;
        }

        int defaultHorizontal() {
            int horizontal;
            if (isDate(clazz) || isDateTime(clazz) || isChar(clazz)) {
                horizontal = Horizontals.CENTER;
            } else if (isInt(clazz) || isLong(clazz)) {
                horizontal = Horizontals.RIGHT;
            } else {
                horizontal = Horizontals.LEFT;
            }
            return horizontal;
        }

        public Column setCellStyle(Font font) {
            this.cellStyle =  styles.of(
                    (font != null ? styles.addFont(font) : 0)
                            | Verticals.CENTER
                            | defaultHorizontal());
            return this;
        }

        public Column setCellStyle(Font font, int horizontal) {
            this.cellStyle =  styles.of(
                    (font != null ? styles.addFont(font) : 0)
                            | Verticals.CENTER
                            | horizontal);
            return this;
        }

        public Column setCellStyle(Font font, Border border) {
            this.cellStyle =  styles.of(
                    (font != null ? styles.addFont(font) : 0)
                            | (border != null ? styles.addBorder(border) : 0)
                            | Verticals.CENTER
                            | defaultHorizontal());
            return this;
        }

        public Column setCellStyle(Font font, Border border, int horizontal) {
            this.cellStyle =  styles.of(
                    (font != null ? styles.addFont(font) : 0)
                            | (border != null ? styles.addBorder(border) : 0)
                            | Verticals.CENTER
                            | horizontal);
            return this;
        }

        public Column setCellStyle(Font font, Fill fill, Border border) {
            this.cellStyle =  styles.of(
                    (font != null ? styles.addFont(font) : 0)
                            | (fill != null ? styles.addFill(fill) : 0)
                            | (border != null ? styles.addBorder(border) : 0)
                            | Verticals.CENTER
                            | defaultHorizontal());
            return this;
        }

        public Column setCellStyle(Font font, Fill fill, Border border, int horizontal) {
            this.cellStyle =  styles.of(
                    (font != null ? styles.addFont(font) : 0)
                            | (fill != null ? styles.addFill(fill) : 0)
                            | (border != null ? styles.addBorder(border) : 0)
                            | Verticals.CENTER
                            | horizontal);
            return this;
        }

        public Column setCellStyle(Font font, Fill fill, Border border, int vertical, int horizontal) {
            this.cellStyle =  styles.of(
                            (font != null ? styles.addFont(font) : 0)
                            | (fill != null ? styles.addFill(fill) : 0)
                            | (border != null ? styles.addBorder(border) : 0)
                            | vertical
                            | horizontal);
            return this;
        }

        public Column setCellStyle(NumFmt numFmt, Font font, Fill fill, Border border, int vertical, int horizontal) {
            this.cellStyle =  styles.of(
                    (numFmt != null ? styles.addNumFmt(numFmt) : 0)
                            | (font != null ? styles.addFont(font) : 0)
                            | (fill != null ? styles.addFill(fill) : 0)
                            | (border != null ? styles.addBorder(border) : 0)
                            | vertical
                            | horizontal);
            return this;
        }

        public Column setShare(boolean share) {
            this.share = share;
            return this;
        }

        protected int getCellStyle(Class clazz) {
            int style;
            if (isString(clazz)) {
                style = Styles.defaultStringBorderStyle();
            } else if (isDate(clazz)) {
                style = Styles.defaultDateBorderStyle();
            } else if (isDateTime(clazz)) {
                style = Styles.defaultTimestampBorderStyle();
            } else if (isInt(clazz) || isLong(clazz)) {
                style = Styles.defaultIntBorderStyle();
                switch (type) {
                    case TYPE_PARENTAGE: // 百分比显示
                        style = Styles.clearNumfmt(style) | styles.addNumFmt(new NumFmt("0%_);[Red]\\(0%\\)"));
                        break;
                    case TYPE_RMB: // 显示人民币
                        style = Styles.clearNumfmt(style) | styles.addNumFmt(new NumFmt("¥0_);[Red]\\(¥0\\)"));
                        break;
                    case TYPE_NORMAL: // 正常显示数字
                        break;
                    default:
                }
            } else if (isFloat(clazz)) {
                style = Styles.defaultDoubleBorderStyle();
                switch (type) {
                    case TYPE_PARENTAGE: // 百分比显示
                        style= Styles.clearNumfmt(style) | styles.addNumFmt(new NumFmt("0.00%_);[Red]\\(0.00%\\)"));
                        break;
                    case TYPE_RMB: // 显示人民币
                        style = Styles.clearNumfmt(style) | styles.addNumFmt(new NumFmt("¥0.00_);[Red]\\(¥0.00\\)"));
                        break;
                    case TYPE_NORMAL: // 正常显示数字
                        break;
                default:
                }
            } else if (isBool(clazz) || isChar(clazz)) {
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

    public Sheet autoSize() {
        this.autoSize = 1;
        return this;
    }

    public Sheet fixSize() {
        this.autoSize = 2;
        return this;
    }

    public Sheet fixSize(double width) {
        this.autoSize = 2;
        for (Column hc : columns) {
            hc.setWidth(width);
        }
        return this;
    }

    public int getAutoSize() {
        return autoSize;
    }

    public String getName() {
        return name;
    }

    public Sheet setName(String name) {
        this.name = name;
        return this;
    }

    public final Column[] getColumns() {
        return columns;
    }

    public Sheet setColumns(final Column[] columns) {
        this.columns = columns.clone();
        for (int i = 0; i < columns.length; i++) {
            columns[i].styles = workbook.getStyles();
        }
        return this;
    }

    public WaterMark getWaterMark() {
        return waterMark;
    }

    public Sheet setWaterMark(WaterMark waterMark) {
        this.waterMark = waterMark;
        return this;
    }

    public boolean isHidden() {
        return hidden;
    }
    public Sheet hidden() {
        this.hidden = true;
        return this;
    }
    /**
     * abstract method close
     */
    public abstract void close();

    public abstract void writeTo(Path xl) throws IOException, ExportException;

    public Sheet addRel(Relationship rel) {
        relManager.add(rel);
        return this;
    }

    protected String getFileName() {
        return "sheet" + id + Const.Suffix.XML;
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

    static boolean isDate(Class<?> clazz) {
        return clazz == java.util.Date.class
                || clazz == java.sql.Date.class
                || clazz == java.time.LocalDate.class;
    }

    static boolean isDateTime(Class<?> clazz) {
        return clazz == java.sql.Timestamp.class
                || clazz == java.time.LocalDateTime.class;
    }

    static boolean isInt(Class<?> clazz) {
        return clazz == int.class || clazz == Integer.class
                || clazz == char.class || clazz == Character.class
                || clazz == byte.class || clazz == Byte.class
                || clazz == short.class || clazz == Short.class;
    }

    static boolean isLong(Class<?> clazz) {
        return clazz == long.class || clazz == Long.class;
    }

    static boolean isFloat(Class<?> clazz) {
        return clazz == double.class || clazz == Double.class
                || clazz == float.class || clazz == Float.class;
    }

    static boolean isBool(Class<?> clazz) {
        return clazz == boolean.class || clazz == Boolean.class;
    }

    static boolean isString(Class<?> clazz) {
        return clazz == String.class || clazz == CharSequence.class;
    }

    static boolean isChar(Class<?> clazz) {
        return clazz == char.class || clazz == Character.class;
    }

    static boolean blockOrDefault(int style) {
        return style == -1 || style == Styles.defaultIntBorderStyle();
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
        buf.append("<sheetFormatPr defaultRowHeight=\"16.5\" baseColWidth=\"");
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
        bw.writeInt(columns.length);
        bw.write("\">");

        int c = 1, defaultStyle = defaultHeadStyle();
        for (Column hc : columns) {
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
        logger.debug("sheet: {} write lines: {}", getName(), rows);
    }

    /**
     * 写行数据
     * @param rs ResultSet
     * @param bw
     */
    protected void writeRow(ResultSet rs, ExtBufferedWriter bw) throws IOException, SQLException {
        // Row number
        int r = ++rows;
        final int len = columns.length;
        bw.write("<row r=\"");
        bw.writeInt(r);
        // default data row height 16.5
        bw.write("\" ht=\"16.5\" spans=\"1:");
        bw.writeInt(len);
        bw.write("\">");

        for (int i = 0; i < len; i++) {
            Column hc = columns[i];

            // t n=numeric (default), s=string, b=boolean, str=function string
            // TODO function <f ca="1" or t="shared" ref="O10:O15" si="0" ... si="10"></f>
            if (isString(hc.clazz)) {
                String s = rs.getString(i + 1);
                writeString(bw, s, i);
            }
            else if (isDate(hc.clazz)) {
                java.sql.Date date = rs.getDate(i + 1);
                writeDate(bw, date, i);
            }
            else if (isDateTime(hc.clazz)) {
                Timestamp ts = rs.getTimestamp(i + 1);
                writeTimestamp(bw, ts, i);
            }
//            else if (isChar(hc.clazz)) {
//                char c = (char) rs.getInt(i + 1);
//                writeChar(bw, c, i);
//            }
            else if (isInt(hc.clazz)) {
                int n = rs.getInt(i + 1);
                writeInt(bw, n, i);
            }
            else if (isLong(hc.clazz)) {
                long l = rs.getLong(i + 1);
                writeLong(bw, l, i);
            }
            else if (isFloat(hc.clazz)) {
                double d = rs.getDouble(i + 1);
                writeDouble(bw, d, i);
            } else if (isBool(hc.clazz)) {
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
        final int len = columns.length;
        bw.write("<row r=\"");
        bw.writeInt(r);
        bw.write("\" ht=\"16.5\" spans=\"1:");
        bw.writeInt(len);
        bw.write("\">");

        for (int i = 0; i < len; i++) {
            Column hc = columns[i];
            // t n=numeric (default), s=string, b=boolean, str=function string
            // TODO function <f ca="1" or t="shared" ref="O10:O15" si="0" ... si="10"></f>
            if (isString(hc.clazz)) {
                String s = rs.getString(i + 1);
                writeStringAutoSize(bw, s, i);
            }
            else if (isDate(hc.clazz)) {
                java.sql.Date date = rs.getDate(i + 1);
                writeDate(bw, date, i);
            }
            else if (isDateTime(hc.clazz)) {
                Timestamp ts = rs.getTimestamp(i + 1);
                writeTimestamp(bw, ts, i);
            }
            else if (isInt(hc.clazz)) {
                int n = rs.getInt(i + 1);
                writeIntAutoSize(bw, n, i);
            }
            else if (isLong(hc.clazz)) {
                long l = rs.getLong(i + 1);
                writeLong(bw, l, i);
            }
            else if (isFloat(hc.clazz)) {
                double d = rs.getDouble(i + 1);
                writeDouble(bw, d, i);
            } else if (isBool(hc.clazz)) {
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

    protected int getStyleIndex(Column hc, Object o) {
        int style = hc.getCellStyle(), styleIndex = hc.styles.of(style);
        if (hc.styleProcessor != null) {
            style = hc.styleProcessor.build(o, style, hc.styles);
            styleIndex = hc.styles.of(style);
        }
        return styleIndex;
    }

    protected void writeString(ExtBufferedWriter bw, String s, int column) throws IOException {
        writeString(bw, s, column, s);
    }

    private void writeString(ExtBufferedWriter bw, String s, int column, Object o) throws IOException {
        Column hc = columns[column];
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
        Column hc = columns[column];
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
        int styleIndex = getStyleIndex(columns[column], o);
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
        int styleIndex = getStyleIndex(columns[column], o);
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
        Column hc = columns[column];
        if (hc.processor == null) {
            writeInt0(bw, n, column);
        } else {
            Object o = hc.processor.conversion(n);
            if (o != null) {
                Class<?> clazz = o.getClass();
                boolean blockOrDefault = blockOrDefault(hc.cellStyle);
                if (isString(clazz)) {
                    if (blockOrDefault) {
                        hc.cellStyle = hc.getCellStyle(String.class);
                    }
                    writeString(bw, o.toString(), column, n);
                }
                else if (isChar(clazz)) {
                    if (blockOrDefault) {
                        hc.cellStyle = Styles.defaultCharBorderStyle();
                    }
                    char c = ((Character) o).charValue();
                    writeChar0(bw, c, column, n);
                }
                else if (isInt(clazz)) {
                    n = ((Integer) o).intValue();
                    writeInt0(bw, n, column, n);
                }
                else if (isLong(clazz)) {
                    long l = ((Long) o).longValue();
                    writeLong(bw, l, column, n);
                }
                else if (isDate(clazz)) {
                    if (blockOrDefault) {
                        hc.cellStyle = hc.getCellStyle(java.util.Date.class);
                    }
                    writeDate(bw, (Date) o, column, n);
                }
                else if (isDateTime(clazz)) {
                    if (blockOrDefault) {
                        hc.cellStyle = hc.getCellStyle(java.sql.Timestamp.class);
                    }
                    writeTimestamp(bw, (Timestamp) o, column, n);
                }
                else if (isFloat(clazz)) {
                    if (blockOrDefault) {
                        hc.cellStyle = hc.getCellStyle(double.class);
                    }
                    writeDouble(bw, ((Double) o).doubleValue(), column, n);
                }
                else if (isBool(clazz)) {
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
        Column hc = columns[column];
        if (hc.processor == null) {
            writeInt0(bw, n, column);
        } else {
            Object o = hc.processor.conversion(n);
            if (o != null) {
                Class<?> clazz = o.getClass();
                boolean blockOrDefault = blockOrDefault(hc.cellStyle);
                if (isString(clazz)) {
                    if (blockOrDefault) {
                        hc.cellStyle = hc.getCellStyle(String.class);
                    }
                    writeStringAutoSize(bw, o.toString(), column, n);
                }
                else if (isChar(clazz)) {
                    if (blockOrDefault) {
                        hc.cellStyle = Styles.defaultCharBorderStyle();
                    }
                    char c = ((Character) o).charValue();
                    writeChar0(bw, c, column, n);
                }
                else if (isInt(clazz)) {
                    int nn = ((Integer) o).intValue();
                    writeInt0(bw, nn, column, n);
                }
                else if (isLong(clazz)) {
                    long l = ((Long) o).longValue();
                    writeLong(bw, l, column, n);
                }
                else if (isDate(clazz)) {
                    if (blockOrDefault) {
                        hc.cellStyle = hc.getCellStyle(java.util.Date.class);
                    }
                    writeDate(bw, (Date) o, column, n);
                }
                else if (isDateTime(clazz)) {
                    if (blockOrDefault) {
                        hc.cellStyle = hc.getCellStyle(java.sql.Timestamp.class);
                    }
                    writeTimestamp(bw, (Timestamp) o, column, n);
                }
                else if (isFloat(clazz)) {
                    if (blockOrDefault) {
                        hc.cellStyle = hc.getCellStyle(double.class);
                    }
                    writeDouble(bw, ((Double) o).doubleValue(), column, n);
                }
                else if (isBool(clazz)) {
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
        Column hc = columns[column];
        if (hc.processor == null) {
            writeChar0(bw, c, column);
        } else {
            Object o = hc.processor.conversion(c);
            if (o != null) {
                Class<?> clazz = o.getClass();
                boolean blockOrDefault = blockOrDefault(hc.cellStyle);
                if (isString(clazz)) {
                    if (blockOrDefault) {
                        hc.cellStyle = hc.getCellStyle(String.class);
                    }
                    writeString(bw, o.toString(), column, c);
                }
                else if (isChar(clazz)) {
                    char cc = ((Character) o).charValue();
                    writeChar0(bw, cc, column, c);
                }
                else if (isInt(clazz)) {
                    if (blockOrDefault) {
                        hc.cellStyle = hc.getCellStyle(int.class);
                    }
                    int n = ((Integer) o).intValue();
                    writeInt0(bw, n, column, c);
                }
                else if (isLong(clazz)) {
                    if (blockOrDefault) {
                        hc.cellStyle = hc.getCellStyle(long.class);
                    }
                    long l = ((Long) o).longValue();
                    writeLong(bw, l, column, c);
                }
                else if (isDate(clazz)) {
                    if (blockOrDefault) {
                        hc.cellStyle = hc.getCellStyle(java.util.Date.class);
                    }
                    writeDate(bw, (Date) o, column, c);
                }
                else if (isDateTime(clazz)) {
                    if (blockOrDefault) {
                        hc.cellStyle = hc.getCellStyle(java.sql.Timestamp.class);
                    }
                    writeTimestamp(bw, (Timestamp) o, column, c);
                }
                else if (isFloat(clazz)) {
                    if (blockOrDefault) {
                        hc.cellStyle = hc.getCellStyle(double.class);
                    }
                    writeDouble(bw, ((Double) o).doubleValue(), column, c);
                }
                else if (isBool(clazz)) {
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
        Column hc = columns[column];
        if (hc.processor == null) {
            writeChar0(bw, c, column);
        } else {
            Object o = hc.processor.conversion(c);
            if (o != null) {
                Class<?> clazz = o.getClass();
                boolean blockOrDefault = blockOrDefault(hc.cellStyle);
                if (isString(clazz)) {
                    if (blockOrDefault) {
                        hc.cellStyle = hc.getCellStyle(String.class);
                    }
                    writeStringAutoSize(bw, o.toString(), column);
                }
                else if (isChar(clazz)) {
                    char cc = ((Character) o).charValue();
                    writeChar0(bw, cc, column, c);
                }
                else if (isInt(clazz)) {
                    if (blockOrDefault) {
                        hc.cellStyle = hc.getCellStyle(int.class);
                    }
                    int n = ((Integer) o).intValue();
                    writeInt0(bw, n, column, c);
                }
                else if (isLong(clazz)) {
                    if (blockOrDefault) {
                        hc.cellStyle = hc.getCellStyle(long.class);
                    }
                    long l = ((Long) o).longValue();
                    writeLong(bw, l, column);
                }
                else if (isDate(clazz)) {
                    if (blockOrDefault) {
                        hc.cellStyle = hc.getCellStyle(java.util.Date.class);
                    }
                    writeDate(bw, (Date) o, column);
                }
                else if (isDateTime(clazz)) {
                    if (blockOrDefault) {
                        hc.cellStyle = hc.getCellStyle(java.sql.Timestamp.class);
                    }
                    writeTimestamp(bw, (Timestamp) o, column);
                }
                else if (isFloat(clazz)) {
                    if (blockOrDefault) {
                        hc.cellStyle = hc.getCellStyle(double.class);
                    }
                    writeDouble(bw, ((Double) o).doubleValue(), column);
                }
                else if (isBool(clazz)) {
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
        int styleIndex = getStyleIndex(columns[column], o);
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
        int styleIndex = getStyleIndex(columns[column], o);
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
        int styleIndex = getStyleIndex(columns[column], o);
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
        int styleIndex = getStyleIndex(columns[column], o);
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
        int styleIndex = getStyleIndex(columns[column], o);
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
        int styleIndex = getStyleIndex(columns[column], null);
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
        final int len = columns.length;
        bw.write("<row r=\"");
        bw.writeInt(r);
        bw.write("\" ht=\"16.5\" spans=\"1:");
        bw.writeInt(len);
        bw.write("\">");

        Styles styles = workbook.getStyles();
        for (int i = 1; i <= len; i++) {
            Column hc = columns[i - 1];
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
                char[] chars = int2Col(columns.length);
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
            for (Column hc : columns) {
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
                    if (isString(hc.clazz)) {
                        if (hc.o == null) {
                            len = 0;
                        } else {
                            len = (int) hc.o;
                        }
//                        len = hc.o.toString().getBytes("GB2312").length;
                    }
                    else if (isDate(hc.clazz)) {
                        len = 10;
                    }
                    else if (hc.clazz == java.sql.Timestamp.class) {
                        len = 20;
                    }
                    else if (isInt(hc.clazz)) {
                        // TODO 根据numFmt计算字符宽度
                        len = hc.type > 0 ? 12 :  11;
                    }
                    else if (isFloat(hc.clazz)) {
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

package net.cua.export.entity.e7;

import net.cua.export.manager.Const;
import net.cua.export.util.FileUtil;
import org.apache.log4j.Logger;
import org.dom4j.Attribute;
import org.dom4j.Document;
import org.dom4j.DocumentException;
import org.dom4j.Element;
import org.dom4j.io.SAXReader;

import java.io.IOException;
import java.lang.reflect.Field;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.*;

/**
 * Created by guanquan.wang at 2018-02-26 13:45
 */
public abstract class AbstractTemplate {
    protected Logger logger = Logger.getLogger(this.getClass().getName());
    static final String inlineStr = "inlineStr";

    Path zipPath;
    Map<String, String> map;
    public AbstractTemplate(Path zipPath) {
        this.zipPath = zipPath;
    }

    /**
     * 文件合法性检查
     * @return
     */
    public boolean check() {
        // Integrity check
        Path contentTypePath = zipPath.resolve("[Content_Types].xml");
        SAXReader reader = new SAXReader();
        Document document;
        try {
            document = reader.read(Files.newInputStream(contentTypePath));
        } catch (DocumentException | IOException e) {
            logger.error("[Content_Types].xml load failed.", e);
            return false;
        }

        List<ContentType.Override> overrides = new ArrayList<>();
        List<ContentType.Default> defaults = new ArrayList<>();
        Iterator<Element> it = document.getRootElement().elementIterator();
        while (it.hasNext()) {
            Element e = it.next();
            if ("Override".equals(e.getName())) {
                overrides.add(new ContentType.Override(e.attributeValue("ContentType"), e.attributeValue("PartName")));
            } else if ("Default".equals(e.getName())) {
                defaults.add(new ContentType.Default(e.attributeValue("ContentType"), e.attributeValue("Extension")));
            }
        }

        return checkDefault(defaults) && checkOverride(overrides);
    }

    protected boolean checkDefault(List<ContentType.Default> list) {
        // Double check
        if (list.isEmpty() || !checkDouble(list)) {
            logger.warn("Default empty or double properties.");
        }
        return true;
    }

    protected boolean checkOverride(List<ContentType.Override> list) {
        // Double check
        if (list.isEmpty() || !checkDouble(list)) {
            logger.warn("Override empty or double properties.");
        }
        // File exists check
        for (ContentType.Override o : list) {
            Path subPath = zipPath.resolve(o.getPartName().substring(1));
            if (!Files.exists(subPath)) {
                logger.error("FileNotFoundException: " + subPath);
                return false;
            }
        }

        return true;
    }

    private boolean checkDouble(List<? extends ContentType.Type> list) {
        list.sort(Comparator.comparing(ContentType.Type::getKey));
        int i = 0, len = list.size() - 1;
        boolean boo = false;
        for (; i < len; i++) {
            if (boo = list.get(i).getKey().equals(list.get(i + 1).getKey()))
                break;
        }
        return !(i < len || boo);
    }

    public void bind(Object o) {
        if (o != null) {
            // Translate object to string hashMap
            map = new HashMap<>();
            if (o instanceof Map) {
                Map<?, ?> map1 = (Map) o;
                map1.forEach((k, v) -> {
                    String vs = v != null ? v.toString() : "";
                    map.put(k.toString(), vs);
                });
            } else {
                Field[] fields = o.getClass().getDeclaredFields();
                for (Field field : fields) {
                    field.setAccessible(true);
                    String value;
                    try {
                        Object v = field.get(o);
                        if (v != null) {
                            value = v.toString();
                        } else value = "";
                    } catch (IllegalAccessException e) {
                        value = "";
                    }
                    map.put(field.getName(), value);
                }
            }
        }
        // Search SharedStrings
        int n1 = bindSstData();
        // inner text
        int n2 = bindSheetData();

        logger.info("Search " + (n1 + n2) + " bind words.");
    }

    protected int bindSstData() {
        Path shareStringPath = zipPath.resolve("xl/sharedStrings.xml");
        SAXReader reader = new SAXReader();
        Document document;
        try {
            document = reader.read(Files.newInputStream(shareStringPath));
        } catch (DocumentException | IOException e) {
            // read style file fail.
            logger.error("shareStrings.xml load failed.", e);
            return 0;
        }

        Element sst = document.getRootElement();
        Attribute countAttr = sst.attribute("count");
        // Empty string
        if (countAttr == null || "0".equals(countAttr.getValue())) {
            return 0;
        }
        int n = 0;
        Iterator<Element> iterator = sst.elementIterator();
        while (iterator.hasNext()) {
            Element si = iterator.next(), t = si.element("t");
            String txt = t.getText();
            if (isPlaceholder(txt)) { // 判断是否是占位符
                // 如果是占位符则对值进行替换
                t.setText(getValue(txt));
                n++;
            }
        }

        if (n > 0) {
            try {
                FileUtil.writeToDisk(document, shareStringPath);
            } catch (IOException e) {
                logger.warn("Rewrite " + shareStringPath + " failed.", e);
                // Do nothing
            }
        }
        return n;
    }

    protected int bindSheetData() {
        // Read content
        Path contentTypePath = zipPath.resolve("[Content_Types].xml");
        SAXReader reader = new SAXReader();
        Document document;
        try {
            document = reader.read(Files.newInputStream(contentTypePath));
        } catch (DocumentException | IOException e) {
            // read style file fail.
            logger.error("shareStrings.xml load failed.", e);
            return 0;
        }

        // Find Override
        List<Element> overrides = document.getRootElement().elements("Override");
        int[] result = overrides.stream()
                .filter(e -> Const.ContentType.SHEET.equals(e.attributeValue("ContentType")))
                .map(e -> zipPath.resolve(e.attributeValue("PartName").substring(1)))
                .mapToInt(this::bindSheet)
                .toArray();

        int n = 0;
        for (int i : result) n += i;
        return n;
    }

    int bindSheet(Path sheetPath) {
        SAXReader reader = new SAXReader();
        Document document;
        try {
            document = reader.read(Files.newInputStream(sheetPath));
        } catch (DocumentException | IOException e) {
            // read style file fail.
            logger.error("shareStrings.xml load failed.", e);
            return 0;
        }

        int n = 0;
        Element sheetData = document.getRootElement().element("sheetData");
        // Each rows
        Iterator<Element> iterator = sheetData.elementIterator();
        while (iterator.hasNext()) {
            Element row = iterator.next();
            // Each cells
            Iterator<Element> it = row.elementIterator();
            while (it.hasNext()) {
                Element cell = it.next();
                Attribute t = cell.attribute("t");
                if (t != null && inlineStr.equals(t.getValue())) {
                    Element tv = cell.element("is").element("t");
                    String txt = tv.getText();
                    if (isPlaceholder(txt)) { // 判断是否是占位符
                        // 如果是占位符则对值进行替换
                        tv.setText(getValue(txt));
                        n++;
                    }
                }
            }
        }

        if (n > 0) {
            try {
                FileUtil.writeToDiskNoFormat(document, sheetPath);
            } catch (IOException e) {
                logger.warn("Rewrite " + sheetData + " failed.", e);
                // Do nothing
            }
        }

        return n;
    }

    ////////////////////////////Abstract function/////////////////////////////

    /**
     * 判断是否包含掩码
     * @param txt
     * @return
     */
    protected abstract boolean isPlaceholder(String txt);

    /**
     * 替换掩码
     * @param txt
     * @return
     */
    protected abstract String getValue(String txt);
}

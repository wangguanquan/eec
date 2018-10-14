package net.cua.excel.entity.e7;

import net.cua.excel.manager.RelManager;
import net.cua.excel.annotation.TopNS;
import net.cua.excel.util.FileUtil;
import net.cua.excel.util.StringUtil;
import org.dom4j.Document;
import org.dom4j.DocumentFactory;
import org.dom4j.Element;

import java.io.IOException;
import java.lang.reflect.Field;
import java.nio.file.Path;
import java.util.*;

/**
 * Created by guanquan.wang at 2017/10/10.
 */
@TopNS(prefix = "", value = "Types", uri = "http://schemas.openxmlformats.org/package/2006/content-types")
public class ContentType {
    private Set<? super Type> set;
    private RelManager relManager;

    public ContentType() {
        set = new HashSet<>();
        relManager = new RelManager();
    }

    public void addRel(Relationship rel) {
        relManager.add(rel);
    }

    static abstract class Type {
        protected String contentType;

        public String getContentType() {
            return contentType;
        }

        public void setContentType(String contentType) {
            this.contentType = contentType;
        }

        public abstract String getKey();
    }

    public static class Default extends Type {
        String extension;

        public Default(String contentType, String extension) {
            this.extension = extension;
            this.contentType = contentType;
        }

        public String getExtension() {
            return extension;
        }

        public void setExtension(String extension) {
            this.extension = extension;
        }

        @java.lang.Override
        public int hashCode() {
            return extension.hashCode();
        }

        @java.lang.Override
        public boolean equals(Object o) {
            if (o == null || !(o instanceof Default)) return false;
            return this == o || extension.equals(((Default)o).extension);
        }

        @java.lang.Override
        public String getKey() {
            return extension;
        }
    }

    public static class Override extends Type {
        String partName;

        public Override(String contentType, String partName) {
            this.partName = partName;
            this.contentType = contentType;
        }

        public String getPartName() {
            return partName;
        }

        public void setPartName(String partName) {
            this.partName = partName;
        }
        @java.lang.Override
        public int hashCode() {
            return partName.hashCode();
        }

        @java.lang.Override
        public boolean equals(Object o) {
            if (o == null || !(o instanceof Override)) return false;
            return this == o || partName.equals(((Override)o).partName);
        }

        @java.lang.Override
        public String getKey() {
            return partName;
        }
    }

    public void add(Type type) {
        set.add(type);
    }

    public void write(Path root) throws IOException {
        // relationship
        relManager.write(root, null);
        // write self
        TopNS topNS = this.getClass().getAnnotation(TopNS.class);
        DocumentFactory factory = DocumentFactory.getInstance();
        //use the factory to create a root element
        Element rootElement = factory.createElement(topNS.value(), topNS.uri()[0]);

        for (Iterator<? super Type> it = set.iterator(); it.hasNext(); ) {
            Object o = it.next();
            Class<?> clazz = o.getClass();
            Element ele = rootElement.addElement(clazz.getSimpleName());
            Field[] fields = clazz.getDeclaredFields()
                    , sfilds = clazz.getSuperclass().getDeclaredFields();
            Field[] newFields = Arrays.copyOf(fields, fields.length + sfilds.length);
            for (int j = fields.length; j < newFields.length; j++) {
                newFields[j] = sfilds[j - fields.length];
            }
            for (Field field : newFields) {
                field.setAccessible(true);
                Class<?> _clazz = field.getType();
                if (_clazz == this.getClass()) {
                    continue;
                }
                try {
                    Object oo = field.get(o);
                    if (oo != null) {
                        ele.addAttribute(StringUtil.uppFirstKey(field.getName()), oo.toString());
                    }
                } catch (IllegalAccessException e) {
                    e.printStackTrace();
                }
            }
        }
        Document doc = factory.createDocument(rootElement);
        FileUtil.writeToDisk(doc, root.resolve("[Content_Types].xml")); // write to desk
    }


}

package net.cua.export.entity.e7;

import net.cua.export.annotation.TopNS;
import net.cua.export.manager.RelManager;
import net.cua.export.util.FileUtil;
import net.cua.export.util.StringUtil;
import org.dom4j.Document;
import org.dom4j.DocumentFactory;
import org.dom4j.Element;

import java.io.File;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

/**
 * Created by wanggq on 2017/10/10.
 */
@TopNS(prefix = "", value = "Types", uri = "http://schemas.openxmlformats.org/package/2006/content-types")
public class ContentType {
    private List<? super Type> list;
    private RelManager relManager;

    public ContentType() {
        list = new ArrayList<>();
        relManager = new RelManager();
    }

    public void addRel(Relationship rel) {
        relManager.add(rel);
    }

    private static class Type {
        protected String contentType;

        public String getContentType() {
            return contentType;
        }

        public void setContentType(String contentType) {
            this.contentType = contentType;
        }
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
    }

    public void add(Type type) {
        list.add(type);
    }

    public void wirte(File root) {
        // relationship
        try {
            relManager.write(root, null);
        } catch (IllegalAccessException e) {
            e.printStackTrace();
        }
        // write self
        TopNS topNS = this.getClass().getAnnotation(TopNS.class);
        DocumentFactory factory = DocumentFactory.getInstance();
        //use the factory to create a root element
        Element rootElement = factory.createElement(topNS.value(), topNS.uri()[0]);

        for (int i = 0, size = list.size(); i < size; i++) {
            Object o = list.get(i);
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
        FileUtil.writeToDisk(doc, root.getPath() + "/[Content_Types].xml"); // write to desk
    }


}

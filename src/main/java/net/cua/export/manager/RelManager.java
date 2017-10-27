package net.cua.export.manager;

import net.cua.export.annotation.TopNS;
import net.cua.export.entity.e7.Relationship;
import net.cua.export.util.FileUtil;
import net.cua.export.util.StringUtil;
import org.dom4j.Document;
import org.dom4j.DocumentFactory;
import org.dom4j.Element;

import java.io.*;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;

/**
 * Created by wanggq on 2017/9/30.
 */
@TopNS(prefix = "", value = "Relationships", uri = "http://schemas.openxmlformats.org/package/2006/relationships")
public class RelManager {
    List<Relationship> relationships;

    public void add(Relationship rel) {
        if (relationships == null) {
            relationships = new ArrayList<>();
        }
        int n = indexOf(rel.getTarget());
        if (n > -1) { // 替换
            rel.setId("rId" + (n + 1));
            relationships.add(n, rel);
        } else { // 追加
            rel.setId("rId" + (relationships.size() + 1));
            relationships.add(rel);
        }
    }

    private int indexOf(String target) {
        if (relationships == null || relationships.isEmpty())
            return -1;
        int i = 0;
        for (Relationship rel : relationships) {
            if (rel.getTarget().equals(target)) {
                return i;
            }
            i++;
        }
        return -1;
    }

    public Relationship getByTarget(String target) {
        int n = indexOf(target);
        return n == -1 ? null : relationships.get(n);
    }

    public Relationship likeByTarget(String target) {
        if (relationships == null || relationships.isEmpty())
            return null;
        for (Relationship rel : relationships) {
            if (rel.getTarget().contains(target)) {
                return rel;
            }
        }
        return null;
    }

    public void write(File parent, String name) throws IllegalAccessException {
        if (relationships == null || relationships.isEmpty()) {
            return;
        }

        File _rels = new File(parent, "_rels");
        if (!_rels.exists()) {
            _rels.mkdirs();
        }
        if (name == null || name.isEmpty()) {
            name = ".rels";
        } else {
            name = name + ".rels";
        }

        TopNS topNS = this.getClass().getAnnotation(TopNS.class);
        DocumentFactory factory = DocumentFactory.getInstance();
        //use the factory to create a root element
        Element rootElement = factory.createElement(topNS.value(), topNS.uri()[0]);

        for (Relationship rel : relationships) {
            Class clazz = rel.getClass();
            String className = clazz.getSimpleName();
            Element ele = rootElement.addElement(className);
            Field[] fields = clazz.getDeclaredFields();
            for (Field field : fields) {
                field.setAccessible(true);
                Object oo = field.get(rel);
                if (oo == null) continue;
                Class _clazz = field.getType();
                if (_clazz == this.getClass()) {
                    continue;
                }
                ele.addAttribute(StringUtil.uppFirstKey(field.getName()), oo.toString());
            }
//            ele.addAttribute("Id", rel.getId())
//                    .addAttribute("Type", rel.getType())
//                    .addAttribute("Target", rel.getTarget())
//                    ;
        }
        Document doc = factory.createDocument(rootElement);
        FileUtil.writeToDisk(doc, _rels.getPath() + "/" + name); // write to desk
    }

}

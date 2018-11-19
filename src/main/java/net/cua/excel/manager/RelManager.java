package net.cua.excel.manager;

import net.cua.excel.annotation.TopNS;
import net.cua.excel.entity.e7.Relationship;
import net.cua.excel.util.FileUtil;
import net.cua.excel.util.StringUtil;
import org.dom4j.Document;
import org.dom4j.DocumentFactory;
import org.dom4j.Element;

import java.io.*;
import java.lang.reflect.Field;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;

/**
 * Relation manager
 * Created by guanquan.wang at 2017/9/30.
 */
@TopNS(prefix = "", value = "Relationships", uri = "http://schemas.openxmlformats.org/package/2006/relationships")
public class RelManager implements Serializable, Cloneable {
    private List<Relationship> relationships;

    public static RelManager of(Relationship ... relationships) {
        RelManager relManager = new RelManager();
        relManager.relationships = new ArrayList<>();
        for (Relationship r : relationships) {
            relManager.relationships.add(r);
        }
        return relManager;
    }
    public synchronized void add(Relationship rel) {
        if (relationships == null) {
            relationships = new ArrayList<>();
        }
        int n = indexOf(rel.getTarget());
        if (n > -1) { // 替换
            rel.setId("rId" + (n + 1));
            relationships.set(n, rel);
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

    public Relationship getById(String id) {
        if (relationships == null || relationships.isEmpty())
            return null;
        for (Relationship rel : relationships) {
            if (rel.getId().equals(id)) {
                return rel;
            }
        }
        return null;
    }

    public void write(Path parent, String name) throws IOException {
        if (relationships == null || relationships.isEmpty()) {
            return;
        }

        Path rels = parent.resolve("_rels");
        if (!Files.exists(rels)) {
            Files.createDirectory(rels);
        }

        if (name == null || name.isEmpty()) {
            name = Const.Suffix.RELATION;
        } else {
            name += Const.Suffix.RELATION;
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
                Object oo = null;
                try {
                    oo = field.get(rel);
                } catch (IllegalAccessException e) {
                    e.printStackTrace();
                }
                if (oo == null) continue;
                Class _clazz = field.getType();
                if (_clazz == this.getClass()) {
                    continue;
                }
                ele.addAttribute(StringUtil.uppFirstKey(field.getName()), oo.toString());
            }
        }
        Document doc = factory.createDocument(rootElement);
        FileUtil.writeToDiskNoFormat(doc, rels.resolve(name)); // write to desk
    }

    @Override
    public RelManager clone() {
        try {
            ByteArrayOutputStream bos = new ByteArrayOutputStream();
            ObjectOutputStream oos = new ObjectOutputStream(bos);
            oos.writeObject(this);

            ObjectInputStream ois = new ObjectInputStream(new ByteArrayInputStream(bos.toByteArray()));
            return (RelManager) ois.readObject();
        } catch (IOException | ClassNotFoundException e) {
            RelManager rm = new RelManager();
            if (relationships != null) {
                for (Relationship r : relationships) {
                    rm.add(r.clone());
                }
            }
            return rm;
        }
    }
}

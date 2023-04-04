/*
 * Copyright (c) 2017, guanquan.wang@yandex.com All Rights Reserved.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

package org.ttzero.excel.manager;

import org.dom4j.Document;
import org.dom4j.DocumentFactory;
import org.dom4j.Element;
import org.ttzero.excel.entity.Relationship;
import org.ttzero.excel.util.FileUtil;
import org.ttzero.excel.util.StringUtil;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.ObjectInputStream;
import java.io.ObjectOutputStream;
import java.io.Serializable;
import java.lang.reflect.Field;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import static org.ttzero.excel.util.FileUtil.exists;

/**
 * Relation manager
 *
 * @author guanquan.wang on 2017/9/30.
 */
@TopNS(prefix = "", value = "Relationships", uri = "http://schemas.openxmlformats.org/package/2006/relationships")
public class RelManager implements Serializable {

    private static final long serialVersionUID = 1L;
    private List<Relationship> relationships;

    public static RelManager of(Relationship... relationships) {
        RelManager relManager = new RelManager();
        relManager.relationships = new ArrayList<>(Arrays.asList(relationships));
        return relManager;
    }

    public synchronized Relationship add(Relationship rel) {
        if (relationships == null) {
            relationships = new ArrayList<>();
        }
        int n = indexOf(rel.getTarget());
        if (n > -1) { // Replace
            rel.setId("rId" + (n + 1));
            relationships.set(n, rel);
        } else { // Append
            rel.setId("rId" + (relationships.size() + 1));
            relationships.add(rel);
        }
        return rel;
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

    public Relationship getByType(String type) {
        if (relationships == null || relationships.isEmpty() || StringUtil.isEmpty(type))
            return null;
        for (Relationship rel : relationships) {
            if (type.equals(rel.getType())) return rel;
        }
        return null;
    }

    public void write(Path parent, String name) throws IOException {
        if (relationships == null || relationships.isEmpty()) {
            return;
        }

        Path rels = parent.resolve("_rels");
        if (!exists(rels)) {
            FileUtil.mkdir(rels);
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
            Class<?> clazz = rel.getClass();
            String className = clazz.getSimpleName();
            Element ele = rootElement.addElement(className);
            Field[] fields = clazz.getDeclaredFields();
            for (Field field : fields) {
                if ("serialVersionUID".equals(field.getName())) {
                    continue;
                }
                field.setAccessible(true);
                Object oo = null;
                try {
                    oo = field.get(rel);
                } catch (IllegalAccessException e) {
                    e.printStackTrace();
                }
                if (oo == null) continue;
                Class<?> _clazz = field.getType();
                if (_clazz == this.getClass()) {
                    continue;
                }
                ele.addAttribute(StringUtil.uppFirstKey(field.getName()), oo.toString());
            }
        }
        Document doc = factory.createDocument(rootElement);
        FileUtil.writeToDiskNoFormat(doc, rels.resolve(name)); // write to desk
    }

    public RelManager deepClone() {
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

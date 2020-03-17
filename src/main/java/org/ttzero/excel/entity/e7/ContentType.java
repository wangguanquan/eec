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

package org.ttzero.excel.entity.e7;

import org.dom4j.Document;
import org.dom4j.DocumentFactory;
import org.dom4j.Element;
import org.ttzero.excel.annotation.TopNS;
import org.ttzero.excel.entity.Relationship;
import org.ttzero.excel.entity.Storable;
import org.ttzero.excel.manager.RelManager;
import org.ttzero.excel.util.FileUtil;
import org.ttzero.excel.util.StringUtil;

import java.io.IOException;
import java.lang.reflect.Field;
import java.nio.file.Path;
import java.util.Arrays;
import java.util.HashSet;
import java.util.Iterator;
import java.util.Set;

/**
 * @author guanquan.wang on 2017/10/10.
 */
@TopNS(prefix = "", value = "Types", uri = "http://schemas.openxmlformats.org/package/2006/content-types")
public class ContentType implements Storable {
    private Set<? super Type> set;
    private RelManager relManager;

    public ContentType() {
        set = new HashSet<>();
        relManager = new RelManager();
    }

    public void addRel(Relationship rel) {
        relManager.add(rel);
    }

    public static abstract class Type {
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
            if (!(o instanceof Default)) return false;
            return this == o || extension.equals(((Default) o).extension);
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
            if (!(o instanceof Override)) return false;
            return this == o || partName.equals(((Override) o).partName);
        }

        @java.lang.Override
        public String getKey() {
            return partName;
        }
    }

    public void add(Type type) {
        set.add(type);
    }

    @java.lang.Override
    public void writeTo(Path root) throws IOException {
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
        FileUtil.writeToDiskNoFormat(doc, root.resolve("[Content_Types].xml")); // write to desk
    }


}

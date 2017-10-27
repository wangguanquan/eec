package net.cua.export.entity;

import net.cua.export.annotation.Attr;
import net.cua.export.annotation.NS;
import net.cua.export.annotation.TopNS;
import net.cua.export.util.DateUtil;
import net.cua.export.util.FileUtil;
import net.cua.export.util.StringUtil;
import org.dom4j.*;
import org.dom4j.io.OutputFormat;
import org.dom4j.io.XMLWriter;

import java.io.*;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.Collection;
import java.util.Date;
import java.util.Iterator;
import java.util.List;

import static net.cua.export.util.StringUtil.indexOf;
import static net.cua.export.util.StringUtil.uppFirstKey;

/**
 * Created by wanggq on 2017/9/21.
 */
public class XmlEntity {

    String[] prefixs, uris;

    public void writeTo(String path) throws IllegalAccessException, NoSuchMethodException, InvocationTargetException {
        DocumentFactory factory = DocumentFactory.getInstance();
        //use the factory to create a root element
        Element rootElement = null;
        //use the factory to create a new document with the previously created root element
        boolean hasTopNs;
        TopNS topNs = getClass().getAnnotation(TopNS.class);
        if (hasTopNs = getClass().isAnnotationPresent(TopNS.class)) {
            prefixs = topNs.prefix();
            uris = topNs.uri();
            for (int i = 0; i < prefixs.length; i++) {
                if (prefixs[i].length() == 0) { // 创建前缀为空的命名空间
                    rootElement = factory.createElement(topNs.value(), uris[i]);
                    break;
                }
            }
        }
        if (rootElement == null) {
            if (hasTopNs) {
                rootElement = factory.createElement(topNs.value());
            } else {
                // TODO echo error message
                return;
            }
        }

        if (hasTopNs) {
            for (int i = 0; i < prefixs.length; i++) {
                if (prefixs.length > 0) {
                    rootElement.add(Namespace.get(prefixs[i], uris[i]));
                }
            }
        }
        toXML(rootElement, this);

        Document doc = factory.createDocument(rootElement);
        FileUtil.writeToDisk(doc, path); // write to desk
    }

    public void toXML(Element doc, Object o) throws IllegalAccessException, NoSuchMethodException, InvocationTargetException {
        Field[] fields = o.getClass().getDeclaredFields();
        for (Field field : fields) {
            field.setAccessible(true);
            Object oo = field.get(o);
            // TODO skip null value
            if (oo == null) {
                continue;
            }
            Class<?> clazz = field.getType();
            if (clazz == this.getClass()) {
                continue;
            }
            Element element;
            NS ns = field.getAnnotation(NS.class);
            if (ns == null && field.isAnnotationPresent(Attr.class)) {
                ns = field.getAnnotation(Attr.class).namespace();
            }
            if (ns != null) {
                Namespace namespace;
                if (ns.uri().length() == 0) {

                    int n = indexOf(prefixs, ns.value());
                    namespace = Namespace.get(ns.value(), n > -1 ? uris[n] : "");
                } else {
                    namespace = Namespace.get(ns.value(), ns.uri());
                    doc.add(namespace);
                }
                // TODO null value write
                if (oo == null) {
                    element = doc.addElement(QName.get(field.getName(), namespace));
                    writeAttr(field, element);
                    continue;
                }
                if (clazz == Date.class || clazz == java.sql.Date.class) {
                    element = doc.addElement(QName.get(field.getName(), namespace)).addText(DateUtil.toTString((Date) oo));
                } else if (clazz == List.class) {
                    element = doc.addElement(QName.get(field.getName(), namespace));
                    Collection collection = (Collection) oo;
                    if (field.isAnnotationPresent(Attr.class)) {
                        Attr attr = field.getAnnotation(Attr.class);
                        String[] names = attr.name(), values = attr.value();
                        NS subNs = attr.namespace();
                        if (!noNamespace(subNs)) {
                            Namespace ans = Namespace.get(subNs.value(), subNs.uri());
                            doc.add(ans);
                            for (int i = 0, len; i < names.length; i++) {
                                if (values[i].charAt(0) == '#' && values[i].charAt(len = values[i].length() - 1) == '#') {
                                    Method m = oo.getClass().getMethod(values[i].substring(1, len), null);
                                    m.setAccessible(true);
                                    Object vo = m.invoke(oo, null);
                                    element.addAttribute(QName.get(names[i], ans), vo.toString());
                                } else {
                                    element.addAttribute(QName.get(names[i], ans), values[i]);
                                }
                            }
                        } else {
                            for (int i = 0, len; i < names.length; i++) {
                                if (values[i].charAt(0) == '#' && values[i].charAt(len = values[i].length() - 1) == '#') {
                                    Method m = oo.getClass().getMethod(values[i].substring(1, len), null);
                                    m.setAccessible(true);
                                    Object vo = m.invoke(oo, null);
                                    element.addAttribute(names[i], vo.toString());
                                } else {
                                    element.addAttribute(names[i], values[i]);
                                }
                            }
                        }
                        int n = indexOf(names, "baseType");
                        if (n == -1) {
                            writeArrayNoBaseType(element, collection);
                        } else {
                            writeArrayWithBaseType(element, collection, ns.contentUse() ? namespace : null, values[n]);
                        }
                    } else {
                        writeArrayNoBaseType(element, collection);
                    }

                    continue;
                } else {
                    element = doc.addElement(QName.get(field.getName(), namespace)).addText(oo.toString());
                }
            } else {
                // TODO null value write
                if (oo == null) {
                    element = doc.addElement(uppFirstKey(field.getName()));
                    writeAttr(field, element);
                    continue;
                }
                if (clazz == Date.class || clazz == java.sql.Date.class) {
                    element = doc.addElement(uppFirstKey(field.getName())).addText(DateUtil.toTString((Date) oo));
                } else if (clazz == List.class) {
                    element = doc.addElement(uppFirstKey(field.getName()));
                    Collection collection = (Collection) oo;
                    if (field.isAnnotationPresent(Attr.class)) {
                        Attr attr = field.getAnnotation(Attr.class);
                        String[] names = attr.name(), values = attr.value();
                        NS subNs = attr.namespace();
                        if (!noNamespace(subNs)) {
                            Namespace ans = Namespace.get(subNs.value(), subNs.uri());
                            doc.add(ans);
                            for (int i = 0, len; i < names.length; i++) {
                                if (values[i].charAt(0) == '#' && values[i].charAt(len = values[i].length() - 1) == '#') {
                                    Object vo = oo.getClass().getMethod(values[i].substring(1, len), null).invoke(oo, null);
                                    element.addAttribute(QName.get(names[i], ans), vo.toString());
                                } else {
                                    element.addAttribute(QName.get(names[i], ans), values[i]);
                                }
                            }
                        } else {
                            for (int i = 0; i < names.length; i++) {
                                element.addAttribute(names[i], values[i]);
                            }
                        }
                        int n = indexOf(names, "baseType");
                        if (n == -1) {
                            writeArrayNoBaseType(element, collection);
                        } else {
                            writeArrayWithBaseType(element, collection, null, values[n]);
                        }
                    } else {
                        writeArrayNoBaseType(element, collection);
                    }

                    continue;
                } else if (isDeclarClass(clazz)) {
                    element = doc.addElement(uppFirstKey(field.getName()));
                    toXML(element, oo);
                } else {
                    element = doc.addElement(uppFirstKey(field.getName())).addText(oo.toString());
                }
            }

            writeAttr(field, element);
        }
    }

    protected void writeAttr(Field field, Element element) {
        if (field.isAnnotationPresent(Attr.class)) {
            Attr attr = field.getAnnotation(Attr.class);
            String[] names = attr.name(), values = attr.value();
            NS _ns = attr.namespace();
            if (!noNamespace(_ns)) {
                Namespace ans = Namespace.get(_ns.value(), _ns.uri());
//                doc.add(ans);
                for (int i = 0; i < names.length; i++) {
                    element.addAttribute(QName.get(names[i], ans), values[i]);
                }
            } else {
                for (int i = 0; i < names.length; i++) {
                    element.addAttribute(names[i], values[i]);
                }
            }
        }
    }

    protected boolean noNamespace(NS ns) {
        return (ns.value().length() == 0 || "-".equals(ns.value()));
    }

    private void writeArrayNoBaseType(Element element, Collection collection) {
        // TODO not namespace array
        StringBuilder buf = new StringBuilder();
        for (Iterator it = collection.iterator(); it.hasNext(); ) {
            Object node = it.next();
            if (node instanceof NameValue) {
                NameValue nv = (NameValue) node;
                element.addElement(nv.getName()).setText(nv.getValue());
            } else {
//                element.addElement(String.valueOf(++idx)).setText(node.toString());
                buf.append(node).append(',');
            }
        }
        if (buf.length() > 0) {
            buf.deleteCharAt(buf.length() - 1);
            element.setText(buf.toString());
        }
    }

    protected void writeArrayWithBaseType(Element element, Collection collection, Namespace namespace, String baseType) {
        if (namespace != null) {
            for (Iterator it = collection.iterator(); it.hasNext(); ) {
                Object node = it.next();
                if (node instanceof NameValue) {
                    NameValue nv = (NameValue) node;
                    element.addElement(QName.get(baseType, namespace))
                            .addElement(QName.get(nv.getName(), namespace)).setText(nv.getValue());
                } else {
                    element.addElement(QName.get(baseType, namespace)).setText(node.toString());
                }
            }
        } else {
            for (Iterator it = collection.iterator(); it.hasNext(); ) {
                Object node = it.next();
                if (node instanceof NameValue) {
                    NameValue nv = (NameValue) node;
                    element.addElement(baseType).addElement(nv.getName()).setText(nv.getValue());
                } else {
                    element.addElement(baseType).setText(node.toString());
                }
            }
        }
    }

    private boolean isDeclarClass(Class clazz) {
        Class[] declareClasses = getClass().getDeclaredClasses();
        for (Class c : declareClasses) {
            if (c == clazz) {
                return true;
            }
        }
        return false;
    }
}

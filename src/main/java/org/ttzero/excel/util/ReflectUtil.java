/*
 * Copyright (c) 2017-2019, guanquan.wang@yandex.com All Rights Reserved.
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

package org.ttzero.excel.util;

import java.beans.IntrospectionException;
import java.beans.Introspector;
import java.beans.MethodDescriptor;
import java.beans.PropertyDescriptor;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.lang.reflect.Parameter;
import java.util.Arrays;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Map;
import java.util.Set;
import java.util.function.Predicate;

/**
 * @author guanquan.wang at 2019-08-15 21:02
 */
public class ReflectUtil {
    private ReflectUtil() { }

    /**
     * List all declared fields that contains all supper class
     *
     * @param beanClass The bean class to be analyzed.
     * @return all declared fields
     */
    public static Field[] listDeclaredFields(Class<?> beanClass) {
        return listDeclaredFields(beanClass, Object.class);
    }

    /**
     * List all declared fields that contains all supper class
     *
     * @param beanClass The bean class to be analyzed.
     * @return all declared fields
     */
    public static Field[] listDeclaredFieldsUntilJavaPackage(Class<?> beanClass) {
        return listDeclaredFieldsUntilJavaPackage(beanClass, Object.class);
    }

    /**
     * List all declared fields that contains all supper class
     *
     * @param beanClass The bean class to be analyzed.
     * @param stopClass The base class at which to stop the analysis.  Any
     *                  methods/properties/events in the stopClass or in its base classes
     *                  will be ignored in the analysis.
     * @return all declared fields
     */
    public static Field[] listDeclaredFields(Class<?> beanClass, Class<?> stopClass) {
        Field[] fields = beanClass.getDeclaredFields();
        int i = fields.length, last = 0;
        for (; (beanClass = beanClass.getSuperclass()) != stopClass && beanClass != null; ) {
            Field[] subFields = beanClass.getDeclaredFields();
            if (subFields.length > 0) {
                if (subFields.length > last) {
                    Field[] tmp = new Field[fields.length + subFields.length];
                    System.arraycopy(fields, 0, tmp, 0, i);
                    fields = tmp;
                    last = tmp.length - i;
                }
                System.arraycopy(subFields, 0, fields, i, subFields.length);
                i += subFields.length;
                last -= subFields.length;
            }
        }
        return fields;
    }

    /**
     * List all declared fields that contains all supper class
     *
     * @param beanClass The bean class to be analyzed.
     * @param stopClass The base class at which to stop the analysis.  Any
     *                  methods/properties/events in the stopClass or in its base classes
     *                  will be ignored in the analysis.
     * @return all declared fields
     */
    public static Field[] listDeclaredFieldsUntilJavaPackage(Class<?> beanClass, Class<?> stopClass) {
        if (isJavaPackage(beanClass)) return new Field[0];
        Field[] fields = beanClass.getDeclaredFields();
        int i = fields.length, last = 0;
        for (; (beanClass = beanClass.getSuperclass()) != stopClass && beanClass != null && !isJavaPackage(beanClass); ) {
            Field[] subFields = beanClass.getDeclaredFields();
            if (subFields.length > 0) {
                if (subFields.length > last) {
                    Field[] tmp = new Field[fields.length + subFields.length];
                    System.arraycopy(fields, 0, tmp, 0, i);
                    fields = tmp;
                    last = tmp.length - i;
                }
                System.arraycopy(subFields, 0, fields, i, subFields.length);
                i += subFields.length;
                last -= subFields.length;
            }
        }
        return fields;
    }

    /**
     * List all declared fields that contains all supper class
     *
     * @param beanClass The bean class to be analyzed.
     * @param filter A field filter
     * @return all declared fields
     */
    public static Field[] listDeclaredFields(Class<?> beanClass, Predicate<Field> filter) {
        return listDeclaredFields(beanClass, Object.class, filter);
    }

    /**
     * List all declared fields that contains all supper class
     *
     * @param beanClass The bean class to be analyzed.
     * @param filter A field filter
     * @return all declared fields
     */
    public static Field[] listDeclaredFieldsUntilJavaPackage(Class<?> beanClass, Predicate<Field> filter) {
        Field[] fields = listDeclaredFieldsUntilJavaPackage(beanClass, Object.class);

        return filter != null ? fieldFilter(fields, filter) : fields;
    }

    /**
     * List all declared fields that contains all supper class
     *
     * @param beanClass The bean class to be analyzed.
     * @param stopClass The base class at which to stop the analysis.  Any
     *                  methods/properties/events in the stopClass or in its base classes
     *                  will be ignored in the analysis.
     * @param filter A field filter
     * @return all declared fields
     */
    public static Field[] listDeclaredFields(Class<?> beanClass, Class<?> stopClass, Predicate<Field> filter) {
        Field[] fields = listDeclaredFields(beanClass, stopClass);

        return filter != null ? fieldFilter(fields, filter) : fields;
    }

    /**
     * List all declared methods that contains all supper class
     *
     * @param beanClass The bean class to be analyzed.
     * @return all declared method
     * @throws IntrospectionException happens during introspection error
     */
    public static Method[] listDeclaredMethods(Class<?> beanClass)
        throws IntrospectionException {
        return listDeclaredMethods(beanClass, Object.class);
    }

    /**
     * List all declared methods that contains all supper class
     *
     * @param beanClass The bean class to be analyzed.
     * @param filter A method filter
     * @return all declared method
     * @throws IntrospectionException happens during introspection error
     */
    public static Method[] listDeclaredMethods(Class<?> beanClass, Predicate<Method> filter)
        throws IntrospectionException {
        Method[] methods = listDeclaredMethods(beanClass);

        return filter != null ? methodFilter(methods, filter) : methods;
    }

    /**
     * List all declared methods that contains all supper class
     *
     * @param beanClass The bean class to be analyzed.
     * @param stopClass The base class at which to stop the analysis.  Any
     *                  methods/properties/events in the stopClass or in its base classes
     *                  will be ignored in the analysis.
     * @return all declared method
     * @throws IntrospectionException happens during introspection error
     */
    public static Method[] listDeclaredMethods(Class<?> beanClass, Class<?> stopClass)
        throws IntrospectionException {
        if (beanClass == stopClass) return new Method[0];
        MethodDescriptor[] methodDescriptors = Introspector.getBeanInfo(beanClass, stopClass).getMethodDescriptors();
        Method[] allMethods = beanClass.getMethods();
        Method[] methods;
        if (methodDescriptors.length > 0) {
            methods = new Method[methodDescriptors.length];
            for (int i = 0; i < methodDescriptors.length; i++) {
                int index = indexOf(allMethods, methodDescriptors[i].getMethod());
                methods[i] = index >= 0 ? allMethods[index] : methodDescriptors[i].getMethod();
            }
        } else methods = new Method[0];

        return methods;
    }

    /**
     * List all declared methods that contains all supper class
     *
     * @param beanClass The bean class to be analyzed.
     * @param stopClass The base class at which to stop the analysis.  Any
     *                  methods/properties/events in the stopClass or in its base classes
     *                  will be ignored in the analysis.
     * @param filter A method filter
     * @return all declared method
     * @throws IntrospectionException happens during introspection error
     */
    public static Method[] listDeclaredMethods(Class<?> beanClass, Class<?> stopClass, Predicate<Method> filter)
        throws IntrospectionException {
        Method[] methods = listDeclaredMethods(beanClass, stopClass);

        return filter != null ? methodFilter(methods, filter) : methods;
    }

    /**
     * List all declared read methods that contains all supper class
     *
     * @param beanClass The bean class to be analyzed.
     * @return all declared method
     * @throws IntrospectionException happens during introspection error
     */
    public static Method[] listReadMethods(Class<?> beanClass) throws IntrospectionException {
        return listReadMethods(beanClass, Object.class);
    }

    /**
     * List all declared read methods that contains all supper class
     *
     * @param beanClass The bean class to be analyzed.
     * @param stopClass The base class at which to stop the analysis.  Any
     *                  methods/properties/events in the stopClass or in its base classes
     *                  will be ignored in the analysis.
     * @return all declared method
     * @throws IntrospectionException happens during introspection error
     */
    public static Method[] listReadMethods(Class<?> beanClass, Class<?> stopClass)
        throws IntrospectionException {
        Method[] methods = listDeclaredMethods(beanClass, stopClass);

        int n = 0;
        for (int i = 0; i < methods.length; i++) {
            Method method = methods[i];
            if (method.getParameterCount() == 0) {
                Class<?> returnType = method.getReturnType();
                if (returnType != void.class && returnType != Void.class) {
                    methods[n++] = method;
                }
            }
        }

        return n < methods.length ? Arrays.copyOf(methods, n) : methods;
    }

    /**
     * List all declared read methods that contains all supper class
     *
     * @param beanClass The bean class to be analyzed.
     * @param filter A method filter
     * @return all declared method
     * @throws IntrospectionException happens during introspection error
     */
    public static Method[] listReadMethods(Class<?> beanClass, Predicate<Method> filter)
        throws IntrospectionException {
        Method[] methods = listReadMethods(beanClass);

        return filter != null ? methodFilter(methods, filter) : methods;
    }

    /**
     * List all declared read methods that contains all supper class
     *
     * @param beanClass The bean class to be analyzed.
     * @param stopClass The base class at which to stop the analysis.  Any
     *                  methods/properties/events in the stopClass or in its base classes
     *                  will be ignored in the analysis.
     * @param filter A method filter
     * @return all declared method
     * @throws IntrospectionException happens during introspection error
     */
    public static Method[] listReadMethods(Class<?> beanClass, Class<?> stopClass, Predicate<Method> filter)
        throws IntrospectionException {
        Method[] methods = listReadMethods(beanClass, stopClass);

        return filter != null ? methodFilter(methods, filter) : methods;
    }

    /**
     * List all declared read methods that contains all supper class
     *
     * @param beanClass The bean class to be analyzed.
     * @param stopClass The base class at which to stop the analysis.  Any
     *                  methods/properties/events in the stopClass or in its base classes
     *                  will be ignored in the analysis.
     * @return all declared method
     * @throws IntrospectionException happens during introspection error
     */
    public static Map<String, Method> readMethodsMap(Class<?> beanClass, Class<?> stopClass)
        throws IntrospectionException {
        Map<String, Method> tmp = new HashMap<>();
        PropertyDescriptor[] propertyDescriptors = Introspector.getBeanInfo(beanClass, stopClass)
            .getPropertyDescriptors();
        for (PropertyDescriptor pd : propertyDescriptors) {
            Method method = pd.getReadMethod();
            if (method != null) tmp.put(pd.getName(), method);
        }
        return tmp;
    }

    /**
     * List all declared write methods that contains all supper class
     *
     * @param beanClass The bean class to be analyzed.
     * @return all declared method
     * @throws IntrospectionException happens during introspection error
     */
    public static Method[] listWriteMethods(Class<?> beanClass) throws IntrospectionException {
        return listWriteMethods(beanClass, Object.class);
    }

    /**
     * List all declared write methods that contains all supper class
     *
     * @param beanClass The bean class to be analyzed.
     * @param stopClass The base class at which to stop the analysis.  Any
     *                  methods/properties/events in the stopClass or in its base classes
     *                  will be ignored in the analysis.
     * @return all declared method
     * @throws IntrospectionException happens during introspection error
     */
    public static Method[] listWriteMethods(Class<?> beanClass, Class<?> stopClass)
        throws IntrospectionException {
        Method[] methods = listDeclaredMethods(beanClass, stopClass);
        Field[] fields = listDeclaredFields(beanClass, stopClass);
        Set<String> tmp = new HashSet<>();
        for (Field field : fields)
            tmp.add("set" + field.getName().toLowerCase());

        int n = 0;
        for (int i = 0; i < methods.length; i++) {
            Method method = methods[i];
            Class<?> returnType = method.getReturnType();
            if (method.getParameterCount() == 1 && (returnType == void.class || returnType == Void.class)
                && tmp.contains(method.getName().toLowerCase())) {
                methods[n++] = method;
            }
        }

        return n < methods.length ? Arrays.copyOf(methods, n) : methods;
    }

    /**
     * List all declared methods that contains all supper class
     *
     * @param beanClass The bean class to be analyzed.
     * @param filter A method filter
     * @return all declared method
     * @throws IntrospectionException happens during introspection error
     */
    public static Method[] listWriteMethods(Class<?> beanClass, Predicate<Method> filter)
        throws IntrospectionException {
        Method[] methods = listWriteMethods(beanClass);

        return filter != null ? methodFilter(methods, filter) : methods;
    }

    /**
     * List all declared methods that contains all supper class
     *
     * @param beanClass The bean class to be analyzed.
     * @param stopClass The base class at which to stop the analysis.  Any
     *                  methods/properties/events in the stopClass or in its base classes
     *                  will be ignored in the analysis.
     * @param filter A method filter
     * @return all declared method
     * @throws IntrospectionException happens during introspection error
     */
    public static Method[] listWriteMethods(Class<?> beanClass, Class<?> stopClass, Predicate<Method> filter)
        throws IntrospectionException {
        Method[] methods = listWriteMethods(beanClass, stopClass);

        return filter != null ? methodFilter(methods, filter) : methods;
    }

    /**
     * List all declared write methods that contains all supper class
     *
     * @param beanClass The bean class to be analyzed.
     * @param stopClass The base class at which to stop the analysis.  Any
     *                  methods/properties/events in the stopClass or in its base classes
     *                  will be ignored in the analysis.
     * @return all declared method
     * @throws IntrospectionException happens during introspection error
     */
    public static Map<String, Method> writeMethodsMap(Class<?> beanClass, Class<?> stopClass)
        throws IntrospectionException {
        Map<String, Method> tmp = new HashMap<>();
        PropertyDescriptor[] propertyDescriptors = Introspector.getBeanInfo(beanClass, stopClass)
            .getPropertyDescriptors();
        for (PropertyDescriptor pd : propertyDescriptors) {
            Method method = pd.getWriteMethod();
            if (method != null) tmp.put(pd.getName(), method);
        }
        return tmp;
    }

    /**
     * Found source method in methods array witch the source method
     * equals it or the has a same method name and return-type and same parameters
     *
     * @param methods the array methods to be found
     * @param source  the source method
     * @return the index in method array
     */
    public static int indexOf(Method[] methods, Method source) {
        int i = 0;
        for (Method method : methods) {
            if (method == null) {
                i++;
                continue;
            }
            if (method.equals(source)) {
                return i;
            }
            if (method.getName().equals(source.getName())
                && method.getReturnType() == source.getReturnType()
                && parameterDeepEquals(method.getParameters(), source.getParameters())) {
                return i;
            }
            i++;
        }
        return -1;
    }

    // Do Filter
    private static Method[] methodFilter(Method[] methods, Predicate<Method> filter) {
        int n = 0;
        for (int i = 0; i < methods.length; i++) {
            Method method = methods[i];
            if (filter.test(method)) {
                if (i != n) methods[n] = method;
                n++;
            }
        }

        return n < methods.length ? Arrays.copyOf(methods, n) : methods;
    }

    private static Field[] fieldFilter(Field[] fields, Predicate<Field> filter) {
        int n = 0;
        for (int i = 0; i < fields.length; i++) {
            Field field = fields[i];
            field.setAccessible(true);
            if (filter.test(field)) {
                if (i != n) fields[n] = field;
                n++;
            }
        }

        return n < fields.length ? Arrays.copyOf(fields, n) : fields;
    }

    // Parameter deep compare
    private static boolean parameterDeepEquals(Parameter[] aClass, Parameter[] bClass) {
        boolean equals = aClass.length == bClass.length;
        if (equals) {
            for (int i = 0; i < aClass.length; i++) {
                if (!(equals = aClass[i].equals(bClass[i])
                    || aClass[i].getType() == bClass[i].getType())
                ) {
                    break;
                }
            }
        }
        return equals;
    }

    /**
     * 判断一个类是否属于{@code java}或{@code JDK}包
     *
     * @param clazz 要检查的类对象
     * @return {@code true}如果该类属于{@code java.*}或{@code jdk.*}包
     */
    public static boolean isJavaPackage(Class<?> clazz) {
        String packageName = clazz.getPackage().getName();
        return packageName.startsWith("java.") || packageName.startsWith("jdk.");
    }
}

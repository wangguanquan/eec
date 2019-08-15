/*
 * Copyright (c) 2019, guanquan.wang@yandex.com All Rights Reserved.
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
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.Arrays;
import java.util.function.Predicate;

import static jdk.nashorn.internal.objects.Global.println;

/**
 * Create by guanquan.wang at 2019-08-15 21:02
 */
public class ReflectUtil {
    private ReflectUtil() {
    }

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
     * @param stopClass The base class at which to stop the analysis.  Any
     *                  methods/properties/events in the stopClass or in its base classes
     *                  will be ignored in the analysis.
     * @return all declared fields
     */
    public static Field[] listDeclaredFields(Class<?> beanClass, Class<?> stopClass) {
        Field[] fields = beanClass.getDeclaredFields();
        int i = fields.length, last = 0;
        for (; (beanClass = beanClass.getSuperclass()) != stopClass; ) {
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
     * @param predicate A field filter
     * @return all declared fields
     */
    public static Field[] listDeclaredFields(Class<?> beanClass, Predicate<Field> predicate) {
        return listDeclaredFields(beanClass, Object.class, predicate);
    }

    /**
     * List all declared fields that contains all supper class
     *
     * @param beanClass The bean class to be analyzed.
     * @param stopClass The base class at which to stop the analysis.  Any
     *                  methods/properties/events in the stopClass or in its base classes
     *                  will be ignored in the analysis.
     * @param predicate A field filter
     * @return all declared fields
     */
    public static Field[] listDeclaredFields(Class<?> beanClass, Class<?> stopClass, Predicate<Field> predicate) {
        Field[] fields = listDeclaredFields(beanClass, stopClass);

        int n = 0;
        for (int i = 0; i < fields.length; i++) {
            Field field = fields[i];
            if (predicate.test(field) && i != n) {
                fields[n++] = field;
            }
        }

        return n < fields.length ? Arrays.copyOf(fields, n) : fields;
    }

    /**
     * List all declared methods that contains all supper class
     *
     * @param beanClass The bean class to be analyzed.
     * @param stopClass The base class at which to stop the analysis.  Any
     *                  methods/properties/events in the stopClass or in its base classes
     *                  will be ignored in the analysis.
     * @return all declared method
     */
    public static Method[] listDeclaredMethods(Class<?> beanClass, Class<?> stopClass)
        throws IntrospectionException {
        MethodDescriptor[] methodDescriptors = Introspector.getBeanInfo(beanClass, stopClass).getMethodDescriptors();
        Method[] methods;
        if (methodDescriptors.length > 0) {
            methods = new Method[methodDescriptors.length];
            for (int i = 0; i < methodDescriptors.length; i++) {
                methods[i] = methodDescriptors[i].getMethod();
            }
        } else methods = new Method[0];

        return methods;
    }
}

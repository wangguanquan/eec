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

import org.junit.Test;
import org.ttzero.excel.annotation.ExcelColumn;
import org.ttzero.excel.annotation.IgnoreExport;
import org.ttzero.excel.annotation.IgnoreImport;

import java.beans.IntrospectionException;
import java.beans.Introspector;
import java.beans.MethodDescriptor;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.Arrays;

import static org.ttzero.excel.Print.println;
import static org.ttzero.excel.util.ReflectUtil.listReadMethods;

/**
 * @author guanquan.wang at 2019-08-15 21:46
 */
public class ReflectUtilTest {
    @Test public void testListDeclaredFields() {
        Field[] fields = ReflectUtil.listDeclaredFields(A.class);

        assert fields.length == 1;
        assert fields[0].getName().equals("a");
    }

    @Test public void testListDeclaredFields2() {
        Field[] fields = ReflectUtil.listDeclaredFields(B.class);

        assert fields.length == 1;
        assert fields[0].getName().equals("a");
    }

    @Test public void testListDeclaredFields3() {
        Field[] fields = ReflectUtil.listDeclaredFields(C.class);

        assert fields.length == 2;
        assert fields[0].getName().equals("c");
        assert fields[1].getName().equals("a");
    }

    @Test public void testListDeclaredFields4() {
        Field[] fields = ReflectUtil.listDeclaredFields(C.class, A.class);

        assert fields.length == 1;
        assert fields[0].getName().equals("c");
    }

    @Test public void testListDeclaredFields5() {
        Field[] fields = ReflectUtil.listDeclaredFields(C.class
            , field -> field.getAnnotation(ExcelColumn.class) != null);

        assert fields.length == 1;
        assert fields[0].getName().equals("a");
    }

    @Test public void testListDeclaredMethod1() throws IntrospectionException {
        Method[] methods = ReflectUtil.listDeclaredMethods(C.class);

        assert methods.length == 6;
    }

    @Test public void testListDeclaredMethod2() throws IntrospectionException {
        Method[] methods = ReflectUtil.listDeclaredMethods(C.class, B.class);

        assert methods.length == 4;
    }

    @Test public void testListDeclaredMethod3() throws IntrospectionException {
        Method[] methods = ReflectUtil.listDeclaredMethods(C.class, method -> method.getName().startsWith("set"));

        assert methods.length == 2;
    }

    @Test public void testListDeclaredMethod4() throws IntrospectionException {
        Method[] methods = ReflectUtil.listDeclaredMethods(C.class, B.class, method -> method.getName().startsWith("set"));

        assert methods.length == 1;
    }

    @Test public void testListReadMethod1() throws IntrospectionException {
        Method[] methods = ReflectUtil.listReadMethods(C.class);

        assert methods.length == 3;
    }

    @Test public void testListReadMethod2() throws IntrospectionException {
        Method[] methods = ReflectUtil.listReadMethods(C.class, A.class);

        assert methods.length == 2;
    }

    @Test public void testListReadMethod3() throws IntrospectionException {
        Method[] methods = ReflectUtil.listReadMethods(C.class
            , method -> method.getAnnotation(ExcelColumn.class) != null);

        assert methods.length == 1;
    }

    @Test public void testListReadMethod4() throws IntrospectionException {
        Method[] methods = ReflectUtil.listReadMethods(C.class, A.class
            , method -> method.getAnnotation(IgnoreExport.class) != null);

        assert methods.length == 0;
    }

    @Test public void testListWriteMethod1() throws IntrospectionException {
        Method[] methods = ReflectUtil.listWriteMethods(C.class);

        assert methods.length == 2;
    }

    @Test public void testListWriteMethod2() throws IntrospectionException {
        Method[] methods = ReflectUtil.listWriteMethods(C.class, A.class);

        assert methods.length == 1;
    }

    @Test public void testListWriteMethod3() throws IntrospectionException {
        Method[] methods = ReflectUtil.listWriteMethods(C.class
            , method -> method.getAnnotation(ExcelColumn.class) != null);

        assert methods.length == 0;
    }

    @Test public void testListWriteMethod4() throws IntrospectionException {
        Method[] methods = ReflectUtil.listWriteMethods(C.class
            , method -> method.getAnnotation(IgnoreImport.class) != null);

        for (Method method : methods)
            println(method);
        assert methods.length == 1;
    }

    @Test public void test() throws IntrospectionException {
        MethodDescriptor[] methods = Introspector.getBeanInfo(C.class, Object.class).getMethodDescriptors();
        for (MethodDescriptor method : methods) {
            println(method);
            Class<?> returnType = method.getMethod().getReturnType();
            println(returnType);
        }
    }

    @Test public void testRewriteMethod() throws IntrospectionException {
        C c = new C() {
            @Override
            @IgnoreImport
            @ExcelColumn("name")
            public void setC(long c) {
                super.c = c;
            }

            @Override
            @IgnoreExport("CODE")
            @ExcelColumn("CODE")
            public long getC() {
                return 9527L;
            }
        };

        Method[] methods = listReadMethods(c.getClass());
        for (Method method : methods) {
            println(method);
            println(Arrays.toString(method.getAnnotations()));
        }
    }

    public static class A {
        @ExcelColumn
        private int a;

        public int getA() {
            return a;
        }

        @IgnoreImport
        public void setA(int a) {
            this.a = a;
        }
    }

    static class B extends A {
        @Override
        public String toString() {
            return "B <- A";
        }

        public int doIt(int a) {
            return 0;
        }
    }

    public static class C extends B {
        private long c;

        @ExcelColumn
        public long getC() {
            return c;
        }

        public void setC(long c) {
            this.c = c;
        }

        @Override
        public int doIt(int a) {
            return 0;
        }
    }
}

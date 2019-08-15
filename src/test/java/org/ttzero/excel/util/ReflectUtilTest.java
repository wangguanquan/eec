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

import org.junit.Test;
import org.ttzero.excel.annotation.ExcelColumn;

import java.beans.IntrospectionException;
import java.beans.Introspector;
import java.beans.MethodDescriptor;
import java.lang.reflect.Field;
import java.lang.reflect.Method;

import static org.ttzero.excel.Print.println;

/**
 * Create by guanquan.wang at 2019-08-15 21:46
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

    @Test public void test() throws IntrospectionException {
        MethodDescriptor[] methods = Introspector.getBeanInfo(C.class, Object.class).getMethodDescriptors();
        for (MethodDescriptor method : methods) {
            println(method);
            Class<?> returnType = method.getMethod().getReturnType();
            println(returnType);
        }
    }

    public static class A {
        @ExcelColumn
        private int a;

        public int getA() {
            return a;
        }

        public void setA(int a) {
            this.a = a;
        }
    }

    static class B extends A {

    }

    public static class C extends B {
        private long c;

        public long getC() {
            return c;
        }

        public void setC(long c) {
            this.c = c;
        }
    }
}

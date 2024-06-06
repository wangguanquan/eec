///*
// * Copyright (c) 2017, guanquan.wang@yandex.com All Rights Reserved.
// *
// * Licensed under the Apache License, Version 2.0 (the "License");
// * you may not use this file except in compliance with the License.
// * You may obtain a copy of the License at
// *
// *     http://www.apache.org/licenses/LICENSE-2.0
// *
// * Unless required by applicable law or agreed to in writing, software
// * distributed under the License is distributed on an "AS IS" BASIS,
// * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
// * See the License for the specific language governing permissions and
// * limitations under the License.
// */
//
//package org.ttzero.excel.manager;
//
//
//import java.lang.annotation.Documented;
//import java.lang.annotation.ElementType;
//import java.lang.annotation.Retention;
//import java.lang.annotation.RetentionPolicy;
//import java.lang.annotation.Target;
//
///**
// * Xml Attribute
// *
// * @author guanquan.wang on 2017/9/21.
// */
//@Target({ElementType.FIELD})
//@Retention(RetentionPolicy.RUNTIME)
//@Documented
//public @interface Attr {
//    /**
//     * attribute name
//     *
//     * @return the names of attr
//     */
//    String[] name();
//
//    /**
//     * attribute value
//     *
//     * @return the values of attr
//     */
//    String[] value() default {};
//
//    /**
//     * namespace
//     *
//     * @return the xml namespace
//     * '-' if do not have namespace
//     */
//    NS namespace() default @NS("-");
//}

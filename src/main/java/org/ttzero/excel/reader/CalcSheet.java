/*
 * Copyright (c) 2017-2019, guanquan.wang@hotmail.com All Rights Reserved.
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


package org.ttzero.excel.reader;

/**
 * 支持解析公式的工作表，可以通过{@link #asCalcSheet}将普通工作表转为{@code CalcSheet}
 *
 * @author guanquan.wang at 2020-01-11 11:36
 * @deprecated 使用 {@link FullSheet}代替，{@code FullSheet}包含{@code CalcSheet}所有功能
 */
@Deprecated
public interface CalcSheet extends Sheet { }

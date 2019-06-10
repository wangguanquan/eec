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

package cn.ttzero.excel.entity;

import cn.ttzero.excel.service.StudentService;

import java.util.List;

/**
 * Custom data source worksheet, the data source can be
 * micro-services, Mybatis, JPA or any other source. If
 * the data source returns an array of json objects, please
 * convert to an object ArrayList or Map ArrayList, the object
 * ArrayList needs to inherit {@code ListSheet}, the Map ArrayList
 * needs to inherit {@code ListMapSheet} and implement
 * the {@code method} method.
 *
 * If other formats cannot be converted to ArrayList, you
 * need to inherit from the base class {@code Sheet} and implement the
 * {@code resetBlockData} and {@code getHeaderColumns} methods.
 *
 * Create by guanquan.wang at 2019-04-30 15:12
 */
public class CustomizeDataSourceSheet extends ListSheet<ListObjectSheetTest.Student> {

    private StudentService service;

    private int pageNo, limit = 1 << 9;

    /**
     * Do not specify the worksheet name
     * Use the default worksheet name Sheet1,Sheet2...
     */
    public CustomizeDataSourceSheet() {
        this(null);
    }

    /**
     * Specify the worksheet name
     * @param name the worksheet name
     */
    public CustomizeDataSourceSheet(String name) {
        super(name);
        this.service = new StudentService();
    }

    /**
     * This method is used for the worksheet to get the data.
     * This is a data source independent method. You can get data
     * from any data source. Since this method is stateless, you
     * should manage paging or other information in your custom Sheet.
     *
     * The more data you get each time, the faster write speed. You
     * should minimize the database query or network request, but the
     * excessive data will put pressure on the memory. Please balance
     * this value before the speed and memory. You can refer to 2^8 ~ 2^10
     *
     * This method is blocked
     *
     * @return The data output to the worksheet, if a null or empty
     * ArrayList returned, mean that the current worksheet is finished written.
     */
    @Override
    public List<ListObjectSheetTest.Student> more() {
        if (pageNo >= 10) {
            return null;
        }
        return service.getPageData(pageNo++, limit);
    }

    /**
     * The worksheet is written by units of row-block. The default size
     * of a row-block is 32, which means that 32 rows of data are
     * written at a time. If the data is not enough, the {@code more}
     * method will be called to get more data.
     *
     * @return the row-block size
     */
    public int getRowBlockSize() {
        return 256;
    }

}

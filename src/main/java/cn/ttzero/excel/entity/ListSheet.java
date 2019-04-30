/*
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

import static cn.ttzero.excel.manager.Const.ROW_BLOCK_SIZE;

/**
 * Create by guanquan.wang at 2019-04-29 20:56
 */
public abstract class ListSheet extends Sheet {
    protected int start, end;

    /**
     * Constructor worksheet
     */
    public ListSheet() {
        super();
    }

    /**
     * Constructor worksheet
     * @param name the worksheet name
     */
    public ListSheet(String name) {
        super(name);
    }

    /**
     * Constructor worksheet
     * @param name the worksheet name
     * @param columns the header info
     */
    public ListSheet(String name, final Column[] columns) {
        super(name, columns);
    }

    /**
     * Constructor worksheet
     * @param name the worksheet name
     * @param waterMark the water mark
     * @param columns the header info
     */
    public ListSheet(String name, WaterMark waterMark, final Column[] columns) {
        super(name, waterMark, columns);
    }

    /**
     * Returns the end index of row-block
     * @return the end index
     */
    protected int getEndIndex() {
        int end = start + ROW_BLOCK_SIZE;
        return end <= this.end ? end : this.end;
    }

    /**
     * Returns total rows in this worksheet
     * @return -1 if unknown
     */
    @Override
    public int size() {
        return end - start;
    }

    /**
     * Split worksheet data
     */
    protected void paging() {
        int len = dataSize(), limit = sheetWriter.getRowLimit() - 1;
        workbook.what("Total size: " + len);
        // paging
        if (len > limit) {
            int page = len / limit;
            if (len % limit > 0) {
                page++;
            }
            // Insert sub-sheet
            for (int i = 1, index = id, last = page - 1, n; i < page; i++) {
                ListSheet sheet = copy();
                sheet.name = name + " (" + i + ")";
                sheet.start = i * limit;
                sheet.end = (n = (i + 1) * limit) < len ? n : len;
                sheet.shouldClose = i == last;
                workbook.insertSheet(index++, sheet);
            }
            // Reset current index
            start = 0;
            end = limit;
            shouldClose = false;
        } else {
            start = 0;
            end = len;
        }
    }

    /**
     * Paging worksheet
     * @return a copy worksheet
     */
    protected abstract ListSheet copy();

    /**
     * Returns total data size before split
     * @return the total size
     */
    protected abstract int dataSize();
}

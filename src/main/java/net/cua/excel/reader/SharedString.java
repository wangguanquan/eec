package net.cua.excel.reader;

import net.cua.excel.tmap.TIntIntHashMap;

import java.io.BufferedReader;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.*;

/**
 * 读取sharedString数据
 * 对于大文件来说不可能全部加载到内存然后根据下标直接取值，
 * 目前做法是进行分区，将数据分为eden, old, hot三个区域，
 * 新读取的数据放入eden区如果eden区查找不到则匀去old区，最后到hot区查找
 * 如果都没有则按下标重新读取该区块到eden区，原eden区数据复制到old区。
 * 加载两次的区块将被标记，被标记的区块有重复读取时被放入hot区，
 * hot区采用LRU页面置换算法进行淘汰。
 * Create by guanquan.wang at 2018-09-27 14:28
 */
class SharedString {
    private Path sstPath;

    SharedString(Path sstPath) {
        this.sstPath = sstPath;
    }

    private int uniqueCount() throws IOException {
        int off = -1;
        try (BufferedReader br = Files.newBufferedReader(sstPath)) {
            char[] bytes = new char[512];
            int n = br.read(bytes);
            String line = new String(bytes, 0, n);
            // Microsoft Excel
            String uniqueCount = " uniqueCount=";
            int index = line.indexOf(uniqueCount), end = index > 0 ? line.indexOf('"', index += (uniqueCount.length() + 1)) : -1;
            if (end > 0) {
                off = Integer.parseInt(line.substring(index, end));
                // WPS
            } else {
                String count = " count=";
                index = line.indexOf(count);
                end = index > 0 ? line.indexOf('"', index += (count.length() + 1)) : -1;
                if (end > 0) {
                    off = Integer.parseInt(line.substring(index, end));
                }
            }
            vt = line.indexOf("<t>", end);
        }
        return off;
    }

    /** 新加载数据放入此区域 */
    private String[] eden;
    /** eden区未命中时将数据复制到此区域 */
    private String[] old;
    /** 每次加载的数量 */
    private int page = 1024;
    /** 整个文件有多少个字符串 */
    private int max = -1, vt = 0, offsetR = 0;
    /** 新生代offset */
    private int offset_eden = -1;
    /** 老年代offset */
    private int offset_old = -1;
    /** 统计各段被访问次数 */
    private TIntIntHashMap count_area = null;
    /** hot world */
    private Hot hot;
    /** 记录各段的起始位置 */
    private Map<Integer, Long> index_area = null;
    // debug
    int total, total_eden, total_old, total_hot;

    SharedString load() throws IOException {
        if (Files.exists(sstPath)) {
            // Get unique count
            max = uniqueCount();
            // Unknown size or big than page
            int default_cap = 10;
            if (max < 0 || max > page) {
                eden = new String[page];
                old = new String[page];

                if (max > 0 && max / page + 1 > default_cap) default_cap = max / page + 1;
                count_area =  new TIntIntHashMap(default_cap);

                hot = new Hot();
            }
            else {
                eden = new String[max];
            }
            index_area = new HashMap(default_cap);
            index_area.put(0, (long) vt);
        } else {
            max = 0;
        }
        return this;
    }

    /**
     * 根据下标获取字符串
     * @param index index of sharedString
     * @return string
     */
    String get(int index) {
        total++;
        if (index < 0 || max > -1 && max < index) {
            throw new IndexOutOfBoundsException("Index: "+index+", Size: "+max);
        }

        // load first
        if (offset_eden == -1) {
            offset_eden = index / page * page;

            if (vt < 0) vt = 0;

            loadXml0();
            test(index);
        }

        String value;
        if (edenRange(index)) {
            value = eden[index - offset_eden];
            total_eden++;
            return value;
        }

        if (oldRange(index)) {
            value = old[index - offset_old];
            total_old++;
            return value;
        }

        // find in hot cache
        value = hot.get(index);

        // can't find in memory cache
        if (value == null) {
            System.arraycopy(eden, 0, old, 0, page);
            offset_old = offset_eden;
            // reload data
            offset_eden = index / page * page;
            eden[0] = null;
            loadXml0();
            if (eden[0] == null) {
                throw new IndexOutOfBoundsException("Index: "+index+", Size: "+max);
            }
            value = eden[index - offset_eden];
            if (test(index)) {
                hot.push(index, value);
            }
            total_eden++;
        } else {
            total_hot++;
        }

        return value;
    }

    private boolean edenRange(int index) {
        return offset_eden >= 0 && offset_eden <= index && offset_eden + page > index;
    }

    private boolean oldRange(int index) {
        return offset_old >= 0 && offset_old <= index && offset_old + page > index;
    }

    /**
     * LRU2
     * @param index
     * @return
     */
    private boolean test(int index) {
        if (max < page) return false;
        int n = count_area.incrementGet(index / page);
        return n > 1;
    }

    /**
     * 分段加载数据
     * TODO 保留一个主要流
     */
    private void loadXml0() {
        int index = offset_eden / page;
        try (BufferedReader br = Files.newBufferedReader(sstPath)) {

            char[] cb = new char[8192];
            int nChar, length, n = 0, offset = 0;
            long skip = index_area.containsKey(index) ? index_area.get(index) : 0L, skipR = skip;

            // 下一个相邻区域或者已读区域
            if (skip > 0L) {
                br.skip(skip); // 跳过前方N个字符从当前位置直接读取
                // 跳过N个区域读取时预先记录跳过的区域起始位置
            } else {
                if (index_area.isEmpty()) {
                    br.skip(skipR = vt);
                } else {
//                    offsetR = max(index_area.keySet());
                    br.skip(skipR = index_area.get(offsetR));
                }

                while (offsetR < index && (length = br.read(cb, offset, cb.length - offset)) > 0) {
                    nChar = 0;
                    length += offset;
                    int cursor = 0;
                    for (int len = length - 4; nChar < len && n < page; nChar++) {
                        if (cb[nChar] == '<' && cb[nChar + 1] == '/' && cb[nChar + 2] == 't' && cb[nChar + 3] == '>') {
                            n++;
                            cursor = nChar += 4;
                        } else continue;
                        // a page
                        if (n == page) {
                            n = 0;
                            index_area.put(++offsetR, skipR + cursor);
                            if (offsetR >= index) break;
                        }
                    }
                    skipR += cursor;
                    System.arraycopy(cb, cursor, cb, 0, offset = length - cursor);
                }
            }

            // Read eden area data
            n = 0;
            while ((length = br.read(cb, offset, cb.length - offset)) > 0 || offset > 0) {
                length += offset;
                nChar = offset &= 0;
                int cursor = 0, len0 = length - 3, len1 = len0 - 1;
                for ( ; nChar < length && n < page; ) {
                    for (; nChar < len0 && (cb[nChar] != '<' || cb[nChar + 1] != 't' || cb[nChar + 2] != '>'); ++nChar);
                    if (nChar >= len0) break; // Not found
                    cursor = nChar;
                    int a = nChar += 3;
                    for (; nChar < len1 && (cb[nChar] != '<' || cb[nChar + 1] != '/' || cb[nChar + 2] != 't' || cb[nChar + 3] != '>'); ++nChar);
                    if (nChar >= len1) break; // Not found
                    eden[n++] = nChar > a ? new String(cb, a, nChar - a) : null;
                    nChar += 4;
                    cursor = nChar;
                }

                skipR += cursor;
                // A page
                if (n == page) {
                    // Save next start index
                    if (offsetR <= index) {
                        index_area.put(++offsetR, skipR);
                    }
                    break;
                } else if (length < cb.length && nChar == len0) { // EOF
                    if (max == -1) { // Reset totals when unknown size
                        max = offsetR * page + n;
                    }
                    break;
                }

                if (cursor < length) {
                    System.arraycopy(cb, cursor, cb, 0, offset = length - cursor);
                }
            }
        } catch (IOException e) {
            throw new ExcelReadException(e);
        }
    }

//    private int max(Set<Integer> set) {
//        Integer max = 0;
//        for (Iterator<Integer> ite = set.iterator(); ite.hasNext(); ) {
//            Integer i = ite.next();
//            if (max.compareTo(i) < 0) {
//                max = i;
//            }
//        }
//        return max;
//    }
}

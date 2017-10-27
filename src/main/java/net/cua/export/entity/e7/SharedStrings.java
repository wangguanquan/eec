package net.cua.export.entity.e7;

import net.cua.export.annotation.TopNS;
import net.cua.export.manager.Const;
import net.cua.export.util.FileUtil;
import net.cua.export.util.StringUtil;
import org.dom4j.Document;
import org.dom4j.DocumentFactory;
import org.dom4j.Element;

import java.io.File;
import java.util.Arrays;
import java.util.Iterator;
import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.concurrent.atomic.AtomicIntegerArray;

/**
 * 共享字符串同一个workbook线程安全
 * Created by wanggq on 2017/10/10.
 */
@TopNS(prefix = "", value = "sst", uri = Const.SCHEMA_MAIN)
public class SharedStrings {
    // 存储共享字符
    private String[] elements;
    private AtomicInteger index; // workbook所有字符串unique长度
    private AtomicInteger count; // workbook所有字符串的个数(shared属性为true)
    /*
     * 查找用cache, 查找并记录次数
     */
    private ConcurrentHashMap<String, AtomicIntegerArray> cache;
    private static final int MAX_CACHE_SIZE = 1 << 8; // 4096

//    private static class Holder { // 每个workbook一个实例
//        private static final SharedStrings INSTANCE = new SharedStrings();
//    }
//
//    public static final SharedStrings getInstance() {
//        return Holder.INSTANCE;
//    }

    SharedStrings() {
        elements = new String[1 << 8];
        cache = new ConcurrentHashMap<>(MAX_CACHE_SIZE);
        index = new AtomicInteger();
        count = new AtomicInteger();
    }

    int search_count = 0;
    /**
     * TODO 解决高速查找办法
     * @param key
     * @return
     */
    public int get(String key) {
//        increment(); // 刷新时间
        search_count++;
        AtomicIntegerArray aia = cache.get(key);
        if (aia != null) {
//            if (aia.get(0) > 1) {
//                aia.addAndGet(0, -2); // 时间-1
//            }
            aia.decrementAndGet(0);

            return aia.get(1);
        }
        // 最大量限制 淘汰最老的词
        if (cache.size() >= MAX_CACHE_SIZE) {
            Map.Entry<String, AtomicIntegerArray> entry = null;
            Iterator<Map.Entry<String, AtomicIntegerArray>> it;
            for (it = cache.entrySet().iterator(); it.hasNext(); ) {
                Map.Entry<String, AtomicIntegerArray> e =  it.next();
                if (entry == null || entry.getValue().get(0) < e.getValue().get(0)) {
                    entry = e;
                }
            }

            // 这里这需要查找时间最大的一组值即可,不用排序
            aia = entry.getValue();
            aia.set(0, ~search_count); // 时间重置
            cache.remove(entry.getKey());
            cache.put(key, aia);
        } else {
            aia = new AtomicIntegerArray(2);
//            aia.incrementAndGet(0);
            aia.set(0, ~search_count);
        }
        cache.put(key, aia);

        int n = indexOf(key);
        if (n == -1) {
            n = index.getAndIncrement();
//            synchronized (elements) {
                if (n >= elements.length) {
                    elements = Arrays.copyOf(elements, n << 1);
                }
//            }
            elements[n] = key;
        }
        aia.set(1, n);
        return n;
    }

    protected int indexOf(String key) {
        int len;
        if ((len = index.get()) == 0) return -1;
        for (int i = 0; i < len; i++) {
            if (key.equals(elements[i])) {
                return i;
            }
        }
        return -1;
    }

    protected void increment() {
        for (AtomicIntegerArray aia : cache.values()) {
            aia.incrementAndGet(0);
        }
    }

    /**
     * 各worksheet字符串相加
     *
     * @param c
     */
    public void addCount(int c) {
        count.addAndGet(c);
    }

    public void write(File root) {
        TopNS topNS = getClass().getAnnotation(TopNS.class);

        // TODO 数据量大时不使用dom4j输出
        DocumentFactory factory = DocumentFactory.getInstance();
        //use the factory to create a root element
        Element rootElement = factory.createElement(topNS.value(), topNS.uri()[0]);
        int len;
        rootElement.addAttribute("uniqueCount", String.valueOf(len = index.get()));
        rootElement.addAttribute("count", String.valueOf(count.get()));

        for (int i = 0; i < len; i++) {
            rootElement.addElement("si").addElement("t").setText(elements[i]);
        }

        Document doc = factory.createDocument(rootElement);
        FileUtil.writeToDisk(doc, root.getPath() + "/" + StringUtil.lowFirstKey(getClass().getSimpleName() + Const.Suffix.XML)); // write to desk

        // destroy
        destroy();
    }

    /**
     * clear memory
     */
    protected void destroy() {
        for (int i = 0, len = elements.length; i < len; i++) {
            elements[i] = null;
        }
        elements = null; // for GC

        cache.clear();
        cache = null;

        index = null;
        count = null;
    }
}

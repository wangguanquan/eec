package net.cua.excel.entity;

import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.util.Arrays;
import java.util.Locale;
import java.util.Properties;

/**
 * 国际化
 * Create by guanquan.wang at 2018-10-13
 */
public class I18N {
    private Properties pro;
    private final String fn = "message._.properties";
    public I18N() {
        Locale locale = Locale.getDefault();
        pro = new Properties();
        try {
            InputStream is = I18N.class.getClassLoader().getResourceAsStream("I18N/" + fn.replace("_", locale.getLanguage()));
            if (is == null) {
                is = I18N.class.getClassLoader().getResourceAsStream("I18N/" + fn.replace("_", "zh_CN"));
            }
            if (is != null) {
                pro.load(new InputStreamReader(is, "UTF-8"));
            }
        } catch (IOException e) {
            // nothing...
        }
    }

    /**
     * get message by code
     * @param code message code
     * @return I18N string
     */
    public String get(String code) {
        if (pro != null) {
            return pro.getProperty(code);
        }
        return code;
    }

    /**
     * get message by code
     * @param code code
     * @param args args
     * @return I18N string
     */
    public String get(String code, String ... args) {
        if (pro == null) return code;
        String msg = pro.getProperty(code);
        char[] oldValue = msg.toCharArray();
        int[] indexs = search(oldValue);
        if (indexs == null) {
            return msg;
        }
        int len = indexs.length >= args.length ? args.length : indexs.length, size = 0;
        for (int i = 0; i < len; size += args[i++].length());
        StringBuilder buf = new StringBuilder(oldValue.length + size - (len << 1));
        buf.append(oldValue, 0, indexs[0]).append(args[0]);
        for (int i = 1; i < len; i++) {
            buf.append(oldValue, size = indexs[i - 1] + 2, indexs[i] - size).append(args[i]);
        }
        if (indexs[len - 1] + 2 < oldValue.length) {
            buf.append(oldValue, size = indexs[len - 1] + 2, oldValue.length - size);
        }
        return buf.toString();
    }

    private static int[] search(char[] value) {
        int[] indexs = new int[16];
        int n = 0;
        for (int i = 0; i < value.length - 1; i++) {
            if (value[i] == '{' && value[i + 1] == '}') {
                indexs[n++] = i++;
                if (n == indexs.length) {
                    int[] _indexs = new int[indexs.length << 1];
                    System.arraycopy(indexs, 0, _indexs, 0, indexs.length);
                    indexs = _indexs;
                }
            }
        }
        return Arrays.copyOfRange(indexs, 0, n);
    }

}

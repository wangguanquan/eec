package net.cua.excel.util;

/**
 * Created by guanquan.wang on 2017/9/30.
 */
public class StringUtil {
    public static boolean isEmpty(String s) {
        return s == null || s.isEmpty();
    }

    public static boolean isNotEmpty(String s) {
        return s != null && s.length() > 0;
    }

    public static int indexOf(String[] array, String v) {
        if (v != null) {
            for (int i = 0; i < array.length; i++) {
                if (v.equals(array[i])) {
                    return i;
                }
            }
        } else {
            for (int i = 0; i < array.length; i++) {
                if (array[i] == null) {
                    return i;
                }
            }
        }
        return -1;
    }


    public static String uppFirstKey(String key) {
        char first = key.charAt(0);
        if (first >= 97 && first <= 122) {
            char[] _v = key.toCharArray();
            _v[0] -= 32;
            return new String(_v);
        }
        return key;
    }

    public static String lowFirstKey(String key) {
        char first = key.charAt(0);
        if (first >= 65 && first <= 90) {
            char[] _v = key.toCharArray();
            _v[0] += 32;
            return new String(_v);
        }
        return key;
    }

    /**
     * convert string to pascal case string
     * @param name
     * @return
     */
    public static String toPascalCase(String name) {
        if (name.indexOf('_') < 0) return name;
        char[] oldValues = name.toLowerCase().toCharArray();
        final int len = oldValues.length;
        int i = 1, idx = i;
        for (int n = len - 1; i < n; i++) {
            char c = oldValues[i], cc = oldValues[i + 1];
            if (c == '_' && cc >= 'a' && cc <= 'z') {
                oldValues[idx++] = (char) (cc - 32);
                i++;
            }
            else {
                oldValues[idx++] = c;
            }
        }
        if (i < len) oldValues[idx++] = oldValues[i];
        return new String(oldValues, 0, idx);
    }
}

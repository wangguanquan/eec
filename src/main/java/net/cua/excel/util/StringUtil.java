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
     * convert string to hump string
     * @param name
     * @return
     */
    public static String toHump(String name) {
//        final char[] oldValues = name.toCharArray();
//        final int len = oldValues.length;
//        char[] values = new char[len];
//
//        char c = oldValues[0];
//        values[0] = c;
//        for (int i = 1, j = i, n = len - 1; i < n; i++) {
//            //
//        }
        // TODO
        return name;
    }
}

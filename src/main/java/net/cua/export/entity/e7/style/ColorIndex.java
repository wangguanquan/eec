package net.cua.export.entity.e7.style;

import java.awt.*;

/**
 * Reference: resources/ColorIndex.html
 * Created by guanquan.wang at 2018-02-06 14:40
 */
public class ColorIndex {
    static final int[] colors = {
            0,         0,         0,         0,         0,         0,         0,         0,
            -16777216, -1,        -65536,    -16711936, -16776961, -256,      -65281,    -16711681,
            -8388608,  -16744448, -16777088, -8355840,  -8388480,  -16744320, -4144960,  -8355712,
            -6710785,  -6737050,  -52,       -3342337,  -10092442, -32640,    -16750900, -3355393,
            -16777088, -65281,    -256,      -16711681, -8388480,  -8388608,  -16744320, -16776961,
            -16724737, -3342337,  -3342388,  -103,      -6697729,  -26164,    -3368449,  -13159,
            -13408513, -13382452, -6697984,  -13312,    -26368,    -39424,    -10066279, -6908266,
            -16764058, -13395610, -16764160, -13421824, -6737152,  -6737050,  -13421671, -13421773
    };

    public static int get(int index) {
        if (index < 0 || index >= colors.length) return 8;
        return colors[index];
    }

    public static int indexOf(Color color) {
        return indexOf(color.getRGB());
    }

    public static int indexOf(int rgb) {
        int i = 8;
        if (rgb >= 0) return i;
        for ( ; i < colors.length; i++) {
            if (colors[i] == rgb) break;
        }
        return i < colors.length ? i : -1;
    }

    public static String toARGB(Color color) {
        return toARGB(color.getRGB());
    }

    public static String toARGB(int rgb) {
        int n;
        char[] chars = new char[8];
        for (int i = 0; i < 4; i++) {
            n = (rgb >> 8 * (3 - i)) & 0xff;
            if (n <= 0xf) {
                chars[i << 1] = '0';
                chars[(i << 1) + 1] = (char) (n < 0xa ? '0' + n : 'a' + n - 0xa);
            } else {
                Integer.toHexString(n).getChars(0, 2, chars, i << 1);
            }
        }
        for (int i = 0; i < chars.length; i++) {
            if (chars[i] >= 'a' && chars[i] <= 'z') {
                chars[i] -= ' ';
            }
        }
        return new String(chars);
    }
}

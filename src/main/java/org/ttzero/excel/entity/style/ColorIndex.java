/*
 * Copyright (c) 2017-2018, guanquan.wang@yandex.com All Rights Reserved.
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

package org.ttzero.excel.entity.style;

import java.awt.Color;

/**
 * The Build-In Color
 * <p>
 * <a href="/ref/ColorIndex.html">ColorIndex</a>
 *
 * @author guanquan.wang at 2018-02-06 14:40
 */
public final class ColorIndex {
    /** Default indexed color */
    public static final int[] indexedColors = {
        -16777216, -1,        -65536,    -16711936, -16776961, -256,      -65281,    -16711681,
        -16777216, -1,        -65536,    -16711936, -16776961, -256,      -65281,    -16711681,
        -8388608,  -16744448, -16777088, -8355840,  -8388480,  -16744320, -4144960,  -8355712,
        -6710785,  -6737050,  -52,       -3342337,  -10092442, -32640,    -16750900, -3355393,
        -16777088, -65281,    -256,      -16711681, -8388480,  -8388608,  -16744320, -16776961,
        -16724737, -3342337,  -3342388,  -103,      -6697729,  -26164,    -3368449,  -13159,
        -13408513, -13382452, -6697984,  -13312,    -26368,    -39424,    -10066279, -6908266,
        -16764058, -13395610, -16764160, -13421824, -6737152,  -6737050,  -13421671, -13421773
    };

    /** Defined 12 base theme color */
    public static final Color[] themeColors = {
            // lt1, dk1
            new Color(255, 255, 255), new Color(  0,   0,   0),
            // lt2, dk2
            new Color(238, 236, 225), new Color( 31,  73, 125),
            new Color( 79, 129, 189), new Color(192,  80,  77),
            new Color(155, 187,  89), new Color(128, 100, 162),
            new Color( 75, 172, 198), new Color(247, 150,  70),
            new Color(  0,   0, 255), new Color(128,   0, 128)
    };

    public static int indexOf(Color color) {
        return indexOf(color.getRGB());
    }

    public static int indexOf(int rgb) {
        if (rgb >= 0) return -1; // alpha=0
        int i = 8;
        for (; i < indexedColors.length; i++) {
            if (indexedColors[i] == rgb) break;
        }
        return i < indexedColors.length ? i : -1;
    }

    /**
     * to argb string
     * @param color color
     * @return argb string
     */
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

    /**
     * to rgb string
     *
     * @param color color
     * @return rgb string
     */
    public static String toRGB(Color color) {
        return toRGB(color.getRGB());
    }

    public static String toRGB(int rgb) {
        int n;
        char[] chars = new char[6];
        for (int i = 0; i < 3; i++) {
            n = (rgb >> 8 * (2 - i)) & 0xff;
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

    /**
     * 颜色比较，{@code null}与{@code Color.BLACK}等价
     *
     * @param color1 颜色1
     * @param color2 颜色2
     * @return true：颜色相同
     */
    public static boolean colorNullEqualsBlack(Color color1, Color color2) {
        int n = (color1 != null ? 1 : 0) | (color2 != null ? 2 : 0);
        boolean r;
        switch (n) {
            case 3:  r = color1.equals(color2);      break;
            case 2:  r = color2.equals(Color.BLACK); break;
            case 1:  r = color1.equals(Color.BLACK); break;
            default: r = true;
        }
        return r;
    }
}

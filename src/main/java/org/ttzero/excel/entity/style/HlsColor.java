/*
 * Copyright (c) 2017-2023, guanquan.wang@yandex.com All Rights Reserved.
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

import org.ttzero.excel.util.StringUtil;

import java.awt.Color;

/**
 * Hue-Luminance-Saturation Color
 *
 * @author guanquan.wang at 2023-01-02 15:59
 */
public class HlsColor {
    public double a;
    /**
     * Hue
     */
    public double h;
    /**
     * Luminance
     */
    public double l;
    /**
     * Saturation
     */
    public double s;

    /**
     * Convert rgba color to hls Color
     *
     * @param rgbColor rgb Color
     * @return hlsColor
     */
    public static HlsColor rgbToHls(Color rgbColor) {
        HlsColor hlsColor = new HlsColor();
        double r = rgbColor.getRed() / 255.0D, g = rgbColor.getGreen() / 255.0D
            , b = rgbColor.getBlue() / 255.0D, a = rgbColor.getAlpha() / 255.0D;
        double min = Math.min(r, Math.min(g, b));
        double max = Math.max(r, Math.max(g, b));
        double delta = max - min;
        if (max == min) {
            hlsColor.h = 0D;
            hlsColor.s = 0D;
            hlsColor.l = max;
            return hlsColor;
        }
        hlsColor.l = (min + max) / 2.0D;
        if (hlsColor.l < 0.5D) hlsColor.s = delta / (max + min);
        else hlsColor.s = delta / (2.0D - max - min);
        if (r == max) hlsColor.h = (g - b) / delta;
        if (g == max) hlsColor.h = 2.0D + (b - r) / delta;
        if (b == max) hlsColor.h = 4.0D + (r - g) / delta;
        hlsColor.h *= 60D;
        if (hlsColor.h < 0D) hlsColor.h += 360D;
        hlsColor.a = a;
        return hlsColor;
    }

    /**
     * Convert hls color to rgba
     *
     * @param hlsColor hls color
     * @return rgbColor
     */
    public static Color hlsToRgb(HlsColor hlsColor) {
        Color rgbColor;
        if (hlsColor.s == 0D) {
            rgbColor = new Color((int) (hlsColor.l * 255D), (int) (hlsColor.l * 255D), (int) (hlsColor.l * 255D), (int) (hlsColor.a * 255D));
            return rgbColor;
        }
        double t1;
        if (hlsColor.l < 0.5D) t1 = hlsColor.l * (1.0D + hlsColor.s);
        else t1 = hlsColor.l + hlsColor.s - (hlsColor.l * hlsColor.s);
        double t2 = 2.0 * hlsColor.l - t1;
        double h = hlsColor.h / 360D;
        double tR = h + (1.0D / 3.0D);
        double r = setColor(t1, t2, tR);
        double tG = h;
        double g = setColor(t1, t2, tG);
        double tB = h - (1.0D / 3.0D);
        double b = setColor(t1, t2, tB);
        rgbColor = new Color((int) (r * 255D), (int) (g * 255D), (int) (b * 255D), (int) (hlsColor.a * 255D));
        return rgbColor;
    }

    private static double setColor(double t1, double t2, double t3) {
        if (t3 < 0D) t3 += 1.0D;
        if (t3 > 1D) t3 -= 1.0D;
        double color;
        if (6.0D * t3 < 1D) color = t2 + (t1 - t2) * 6.0D * t3;
        else if (2.0D * t3 < 1D) color = t1;
        else if (3.0D * t3 < 2D) color = t2 + (t1 - t2) * ((2.0D / 3.0D) - t3) * 6.0D;
        else color = t2;
        return color;
    }

    public static double calculateFinalLumValue(Double tint, double lum) {
        if (tint == null) return lum;
        if (tint < 0D) lum = lum * (1.0D + tint);
        else lum = lum * (1.0D - tint) + (255D - 255D * (1.0D - tint));
        return lum;
    }

    public static Color calculateColor(String themeV, String tintV) {
        int theme = 0;
        Double tint = null;
        try {
            theme = Integer.parseInt(themeV);
            if (StringUtil.isNotEmpty(tintV)) {
                tint = Double.parseDouble(tintV);
            }
        } catch (NumberFormatException ex) {
            // Ignore
        }

        Color color;
        // Theme value range 0 ... 9
        if (theme >= 0 && theme <= 9) {
            Color base = ColorIndex.themeColors[theme];
            HlsColor hls = rgbToHls(base);
            hls.l = calculateFinalLumValue(tint, hls.l * 255D) / 255D;
            color = hlsToRgb(hls);
        } else color = ColorIndex.themeColors[0];
        return color;
    }
}

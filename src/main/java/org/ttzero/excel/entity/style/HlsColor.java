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
 * Hue-Luminance-Saturation Color Util (OpenXML theme color)
 *
 * @author guanquan.wang at 2023-01-02 15:59
 */
public class HlsColor {
    /**
     * Alpha
     */
    public float a;
    /**
     * Hue
     */
    public float h;
    /**
     * Luminance
     */
    public float l;
    /**
     * Saturation
     */
    public float s;

    /**
     * Convert rgba color to hls Color
     *
     * @param rgbColor rgb Color
     * @return hlsColor
     */
    public static HlsColor rgbToHls(Color rgbColor) {
        HlsColor hlsColor = new HlsColor();
        float r = rgbColor.getRed() / 255.0F, g = rgbColor.getGreen() / 255.0F
            , b = rgbColor.getBlue() / 255.0F, a = rgbColor.getAlpha() / 255.0F;
        float min = Math.min(r, Math.min(g, b));
        float max = Math.max(r, Math.max(g, b));
        float delta = max - min;
        if (max == min) {
            hlsColor.h = 0F;
            hlsColor.s = 0F;
            hlsColor.l = max;
            return hlsColor;
        }
        hlsColor.l = (min + max) / 2.0F;
        if (hlsColor.l < 0.5F) hlsColor.s = delta / (max + min);
        else hlsColor.s = delta / (2.0F - max - min);
        if (r == max) hlsColor.h = (g - b) / delta;
        if (g == max) hlsColor.h = 2.0F + (b - r) / delta;
        if (b == max) hlsColor.h = 4.0F + (r - g) / delta;
        hlsColor.h *= 60F;
        if (hlsColor.h < 0F) hlsColor.h += 360F;
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
        if (hlsColor.s == 0) {
            rgbColor = new Color(hlsColor.l, hlsColor.l, hlsColor.l, hlsColor.a);
            return rgbColor;
        }
        float t1;
        if (hlsColor.l < 0.5F) t1 = hlsColor.l * (1.0F + hlsColor.s);
        else t1 = hlsColor.l + hlsColor.s - (hlsColor.l * hlsColor.s);
        float t2 = 2.0F * hlsColor.l - t1;
        float h = hlsColor.h / 360F;
        float tR = h + (1.0F / 3.0F);
        float r = hueToRGB(t1, t2, tR);
        float tG = h;
        float g = hueToRGB(t1, t2, tG);
        float tB = h - (1.0F / 3.0F);
        float b = hueToRGB(t1, t2, tB);
//        rgbColor = new Color((int) (r * 255), (int) (g * 255), (int) (b * 255), (int) (hlsColor.a * 255));
        rgbColor = new Color(r, g, b, hlsColor.a);
        return rgbColor;
    }

    private static float hueToRGB(float t1, float t2, float t3) {
        if (t3 < 0) t3 += 1.0F;
        if (t3 > 1) t3 -= 1.0F;
        float color;
        if (6.0F * t3 < 1) color = t2 + (t1 - t2) * 6.0F * t3;
        else if (2.0F * t3 < 1) color = t1;
        else if (3.0F * t3 < 2) color = t2 + (t1 - t2) * ((2.0F / 3.0F) - t3) * 6.0F;
        else color = t2;
        return color;
    }

    /**
     * If tint is supplied, then it is applied to the value of the color to determine the final color applied.
     * <p>
     * The tint value is stored as a double from {@code -1.0 .. 1.0}, where {@code -1.0} means {@code 100%} darken
     * and {@code 1.0} means {@code 100%} lighten. Also, {@code 0.0} means no change.
     * <p>
     * In loading the value, it is converted to HLS where HLS values are (0..HLSMAX), where HLSMAX is currently 255.
     * <p>
     * Referer: <a href="https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.color?view=openxml-2.8.1">Tint</a>
     *
     * @param tint Specifies the tint value applied to the color.
     * @param lum Luminance part
     * @return calculate tint Luminance
     */
    public static float calculateFinalLumValue(Double tint, float lum) {
        if (tint == null) return lum;
        if (tint < 0) lum = (float) (lum * (1 + tint));
        else lum = (float) (lum * (1 - tint) + (255 - 255 * (1 - tint)));
        return lum;
    }

    /**
     * Convert RGB to HSL and then adjust the luminance part
     *
     * @param theme rgb color
     * @param tintV tint a double from {@code -1.0 .. 1.0}
     * @return rgb color
     */
    public static Color calculateColor(Color theme, String tintV) {
        Double tint = null;
        if (StringUtil.isNotEmpty(tintV)) {
            try {
                tint = Double.parseDouble(tintV);
            } catch (NumberFormatException ex) {
                // Ignore
            }
        }

        // Theme value range 0 ... 9
        HlsColor hls = rgbToHls(theme);
        hls.l = calculateFinalLumValue(tint, hls.l * 255) / 255;
        return hlsToRgb(hls);
    }
}

/*
 * Copyright (c) 2017, guanquan.wang@yandex.com All Rights Reserved.
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

package org.ttzero.excel.util;

import java.sql.Time;
import java.sql.Timestamp;
import java.text.SimpleDateFormat;
import java.time.Instant;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeFormatterBuilder;
import java.time.format.DateTimeParseException;
import java.util.Date;
import java.util.TimeZone;

import static java.time.format.DateTimeFormatter.ISO_LOCAL_DATE;
import static java.time.format.DateTimeFormatter.ISO_LOCAL_TIME;

/**
 * Excel07日期工具类，Excel07日期从1900年1月1日开始对应数字{@code 1}，每增加一天数字加1所以2表示{@code 1900-1-2}日，
 * 小数部分表示时分秒在1天中的比例，例中午12点它在1天中比例为0.5所以{@code 1.5}就表示{@code 1900-1-1 12:00:00}
 *
 * @author guanquan.wang on 2017/9/21.
 */
public class DateUtil {
    /**
     * Java timestamp从1970年开始，所以这里计算从1900到1970之前相差的天数
     */
    public static final int DAYS_1900_TO_1970 = ~(int) LocalDate.of(1900, 1, 1).toEpochDay() + 3;
    /**
     * 保存1天的秒数
     */
    public static final double SECOND_OF_DAY = 24 * 60 * 60.0D;

    /**
     * 时区，默认随系统，外部可根据实际的数据时区修改此值
     */
    public static final int tz = TimeZone.getDefault().getRawOffset();

    /**
     * 通用日期格式化 {@code yyyy-MM-dd'T'HH:mm:ss'Z'}
     */
    public static final ThreadLocal<SimpleDateFormat> utcDateTimeFormat
        = ThreadLocal.withInitial(() -> new SimpleDateFormat("yyyy-MM-dd\'T\'HH:mm:ss\'Z\'"));
    /**
     * 通用日期格式化 {@code yyyy-MM-dd HH:mm:ss}
     */
    public static final ThreadLocal<SimpleDateFormat> dateTimeFormat
        = ThreadLocal.withInitial(() -> new SimpleDateFormat("yyyy-MM-dd HH:mm:ss"));
    /**
     * 通用日期格式化 {@code yyyy-MM-dd}
     */
    public static final ThreadLocal<SimpleDateFormat> dateFormat
        = ThreadLocal.withInitial(() -> new SimpleDateFormat("yyyy-MM-dd"));
    /**
     * 通用日期格式化 {@code yyyy-MM-dd HH:mm:ss}
     */
    public static final DateTimeFormatter LOCAL_DATE_TIME;

    static {
        LOCAL_DATE_TIME = new DateTimeFormatterBuilder()
            .parseCaseInsensitive()
            .append(ISO_LOCAL_DATE)
            .appendLiteral(' ')
            .append(ISO_LOCAL_TIME)
            .toFormatter();
    }

    /**
     * 将日期{@code date}格式化为{@code yyyy-MM-dd HH:mm:ss}
     *
     * @param date 待转换日期
     * @return {@code yyyy-MM-dd HH:mm:ss}格式字符串
     * @deprecated 使用 {@link #toDateTimeString(Date)} 替代
     */
    @Deprecated
    public static String toString(Date date) {
        return toDateTimeString(date);
    }
    /**
     * 将日期{@code date}格式化为{@code yyyy-MM-dd HH:mm:ss}
     *
     * @param date 待转换日期
     * @return {@code yyyy-MM-dd HH:mm:ss}格式字符串
     */
    public static String toDateTimeString(Date date) {
        return dateTimeFormat.get().format(date);
    }
    /**
     * 将日期{@code date}格式化为{@code yyyy-MM-dd'T'HH:mm:ss'Z'}
     *
     * @param date 待转换日期
     * @return {@code yyyy-MM-dd'T'HH:mm:ss'Z'}格式字符串
     */
    public static String toTString(Date date) {
        return utcDateTimeFormat.get().format(date);
    }
    /**
     * 将日期{@code date}格式化为{@code yyyy-MM-dd}
     *
     * @param date 待转换日期
     * @return {@code yyyy-MM-dd}格式字符串
     */
    public static String toDateString(Date date) {
        return dateFormat.get().format(date);
    }
    /**
     * 将日期{@code date}格式化为{@code HH:mm:ss}
     *
     * @param date 待转换日期
     * @return {@code HH:mm:ss}格式字符串
     */
    public static String toTimeString(Date date) {
        LocalTime lt = new Timestamp(date.getTime()).toLocalDateTime().toLocalTime();
        char[] chars = new char[8];
        timeChars(lt, chars);
        return new String(chars);
    }
    /**
     * 获取当日{@code yyyy-MM-dd}字符串
     *
     * @return 当日格式化为{@code yyyy-MM-dd}格式字符串
     */
    public static String today() {
        return LocalDate.now().toString();
    }

    private DateUtil() { }
    /**
     * 将{@code Timestamp}转为距{@code 1900-1-1}相差的值，精准到秒
     *
     * @param ts java.sql.Timestamp(not null)
     * @return 距{@code 1900-1-1}相差的值
     */
    public static double toDateTimeValue(Timestamp ts) {
        LocalDateTime ldt = ts.toLocalDateTime();
        long day = ldt.toLocalDate().toEpochDay();
        int second = ldt.toLocalTime().toSecondOfDay();
        return second / SECOND_OF_DAY + day + DAYS_1900_TO_1970;
    }

    /**
     * 将{@code Date}转为距{@code 1900-1-1}相差的值，精准到秒
     *
     * @param date java.util.Date(not null)
     * @return 距{@code 1900-1-1}相差的值
     */
    public static double toDateTimeValue(Date date) {
        return toDateTimeValue(new Timestamp(date.getTime()));
    }

    /**
     * 将{@code Date}转为距{@code 1900-1-1}相差的天数，精准到天
     *
     * @param date java.util.Date(not null)
     * @return 距{@code 1900-1-1}相差的天数
     */
    public static int toDateValue(Date date) {
        int n;
        if (date instanceof java.sql.Date) {
            n = (int) LocalDate.parse(toDateString(date)).toEpochDay() + DAYS_1900_TO_1970;
        } else {
            n = (int) date.toInstant().atZone(ZoneId.systemDefault()).toLocalDate().toEpochDay() + DAYS_1900_TO_1970;
        }
        return n;
    }

    /**
     * 取{@code Timestamp}的时分秒转时分秒在一天的比值{@code second-of-day}
     *
     * @param ts java.sql.Timestamp(not null)
     * @return 时分秒在一天的比值{@code second-of-day}
     */
    public static double toTimeValue(Timestamp ts) {
        return toTimeValue(ts.toLocalDateTime().toLocalTime());
    }

    /**
     * 取{@code Date}的时分秒转为时分秒在一天的比值{@code second-of-day}
     *
     * @param date java.util.Date(not null)
     * @return 时分秒在一天的比值{@code second-of-day}
     */
    public static double toTimeValue(Date date) {
        return toTimeValue(new Timestamp(date.getTime()));
    }

    /**
     * 将{@code java.sql.Time}转为时分秒在一天的比值{@code second-of-day}
     *
     * @param time java.sql.Time(not null)
     * @return 时分秒在一天的比值{@code second-of-day}
     */
    public static double toTimeValue(Time time) {
        return toTimeValue(time.toLocalTime());
    }

    /**
     * 将{@code LocalDateTime}转为距{@code 1900-1-1}相差的值，精准到秒
     *
     * @param ldt java.time.LocalDateTime(not null)
     * @return 距{@code 1900-1-1}相差的值
     */
    public static double toDateTimeValue(LocalDateTime ldt) {
        long day = ldt.toLocalDate().toEpochDay();
        int second = ldt.toLocalTime().toSecondOfDay();
        return second / SECOND_OF_DAY + day + DAYS_1900_TO_1970;
    }

    /**
     * 将{@code LocalDate}转为距{@code 1900-1-1}相差的天数，精准到天
     *
     * @param date java.time.LocalDate(not null)
     * @return 距{@code 1900-1-1}相差的天数
     */
    public static int toDateValue(LocalDate date) {
        return (int) date.toEpochDay() + DAYS_1900_TO_1970;
    }

    /**
     * 将{@code java.sql.Time}转为时分秒在一天的比值{@code second-of-day}
     *
     * @param time LocalTime(not null)
     * @return 时分秒在一天的比值 {@code second-of-day}
     */
    public static double toTimeValue(LocalTime time) {
        return time.toSecondOfDay() / SECOND_OF_DAY;
    }

    /////////////////////////////number to date//////////////////////////////////

    /**
     * Excel07时间转为{@code java.util.Date}
     *
     * @param n excel读取的时间值
     * @return java.util.Date
     */
    public static java.util.Date toDate(int n) {
        return Date.from(Instant.ofEpochSecond((n - DAYS_1900_TO_1970) * 86400L).minusMillis(tz));
    }

    /**
     * Excel07时间转为{@code java.util.Date}
     *
     * @param d excel读取的时间值
     * @return java.util.Date
     */
    public static java.util.Date toDate(double d) {
        int n = (int) d, m = (int) ((d - n) * SECOND_OF_DAY + 0.5D); // Causes data over 0.5s to be carried over to 1s
        return Date.from(Instant.ofEpochSecond((n - DAYS_1900_TO_1970) * 86400L + m).minusMillis(tz));
    }

    /**
     * 时间字符串转时间，最少需要包含{@code yyyy-MM-dd}
     *
     * @param dateStr 时间字符串
     * @return java.util.Date
     */
    public static java.util.Date toDate(String dateStr) {
        return new java.util.Date(toTimestamp(dateStr).getTime());
    }

    /**
     * Excel07时间转为{@code java.sql.Time}
     *
     * @param d excel读取的时间值
     * @return java.sql.Time
     */
    public static java.sql.Time toTime(double d) {
        int n = (int) d, m = (int) ((d - n) * SECOND_OF_DAY + 0.5D); // Causes data over 0.5s to be carried over to 1s
        return java.sql.Time.valueOf(LocalTime.ofSecondOfDay(m));
    }

    /**
     * 时间字符串转时间，字符串格式必须为{@code HH:mm:ss}
     *
     * @param s 时间字符串
     * @return java.sql.Time
     */
    public static java.sql.Time toTime(String s) {
        LocalTime time = toLocalTime(s);
        return time != null ? java.sql.Time.valueOf(time) : null;
    }

    /**
     * Excel07时间转为{@code java.time.LocalTime}
     *
     * @param d excel读取的时间值
     * @return java.time.LocalTime
     */
    public static LocalTime toLocalTime(double d) {
        int n = (int) d, m = (int) ((d - n) * SECOND_OF_DAY + 0.5D); // Causes data over 0.5s to be carried over to 1s
        return LocalTime.ofSecondOfDay(m);
    }

    /**
     * 时间字符串转时间，字符串格式必须为{@code HH:mm:ss}
     *
     * @param s 时间字符串
     * @return java.time.LocalTime
     */
    public static LocalTime toLocalTime(String s) {
        try {
            return LocalTime.parse(s);
        } catch (DateTimeParseException | NullPointerException e) {
            try {
                return toTimestamp(s).toLocalDateTime().toLocalTime();
            } catch (Exception ex) {
                throw new NumberFormatException(s);
            }
        }
    }

    /**
     * Excel07时间转为{@code java.sql.Timestamp}
     *
     * @param d excel读取的时间值
     * @return java.sql.Timestamp
     */
    public static java.sql.Timestamp toTimestamp(double d) {
        int n = (int) d, m = (int) ((d - n) * SECOND_OF_DAY + 0.5D); // Causes data over 0.5s to be carried over to 1s
        return Timestamp.from(Instant.ofEpochSecond((n - DAYS_1900_TO_1970) * 86400L + m).minusMillis(tz));
    }

    /**
     * Excel07时间转为{@code java.time.LocalDateTime}
     *
     * @param d excel读取的时间值
     * @return java.time.LocalDateTime
     */
    public static LocalDateTime toLocalDateTime(double d) {
        int n = (int) d, m = (int) ((d - n) * SECOND_OF_DAY + 0.5D); // Causes data over 0.5s to be carried over to 1s
        return LocalDateTime.ofInstant(Instant.ofEpochSecond((n - DAYS_1900_TO_1970) * 86400L + m).minusMillis(tz), ZoneId.systemDefault());
    }
    /**
     * Excel07时间转为{@code java.sql.Timestamp}
     *
     * @param n excel读取的时间值
     * @return java.sql.Timestamp
     */
    public static java.sql.Timestamp toTimestamp(int n) {
        return Timestamp.from(Instant.ofEpochSecond((n - DAYS_1900_TO_1970) * 86400L).minusMillis(tz));
    }

    /**
     * Excel07时间转为{@code java.time.LocalDate}
     *
     * @param n excel读取的时间值
     * @return java.time.LocalDate
     */
    public static LocalDate toLocalDate(int n) {
        return LocalDate.ofEpochDay(n - DAYS_1900_TO_1970);
    }

    /**
     * 时间字符串转时间，最少需要包含{@code yyyy-MM-dd}
     *
     * @param dateStr 时间字符串
     * @return java.sql.Timestamp
     */
    public static java.sql.Timestamp toTimestamp(String dateStr) {
        String v = dateStr.trim().replace('/', '-');
        int dividingSpace = v.indexOf(' ');
        if (dividingSpace < 0) {
            v += " 00:00:00";
        } else {
            int i = 0, idx = dividingSpace;
            for (; (idx = v.indexOf(':', idx + 1)) > 0;i++);
            boolean endOfp = v.charAt(v.length() - 1) == ':';
            switch (i) {
                case 0: v += ":0:0"; break;
                case 1: v += !endOfp ? ":0" : "0:0"; break;
                case 2: if (endOfp) v += '0'; break;
                default:
            }
        }
        return java.sql.Timestamp.valueOf(v);
    }

    /**
     * 将给定的 LocalTime 对象转换为指定格式的字符数组
     *
     * <p>注意：内部使用，外部勿用</p>
     *
     * @param time 要转换的 LocalTime 对象
     * @param chars 用于存储转换结果的字符数组，长度固定{@code 8}
     * @return 转换后的字符数组。
     */
    public static char[] timeChars(LocalTime time, char[] chars) {
        int hms = time.getHour() * 10000 + time.getMinute() * 100 + time.getSecond();
        for (int i = chars.length - 1; i >= 0; chars[i--] = (char) (hms % 10 + '0'), hms /= 10) {
            if (i == 5 || i == 2) chars[i--] = ':';
        }
        return chars;
    }
}

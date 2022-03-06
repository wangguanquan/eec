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

import org.ttzero.excel.reader.UncheckedTypeException;

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
 * A date util ror excel
 *
 * @author guanquan.wang on 2017/9/21.
 */
public class DateUtil {
    public static final int DAYS_1900_TO_1970 = ~(int) LocalDate.of(1900, 1, 1).toEpochDay() + 3;
    public static final double SECOND_OF_DAY = (double) 24 * 60 * 60;
    public static final double MILLIS_OF_DAY = (double) 24 * 60 * 60 * 1000;

    // time-zone
    public static final int tz = TimeZone.getDefault().getRawOffset();

    public static final ThreadLocal<SimpleDateFormat> utcDateTimeFormat
        = ThreadLocal.withInitial(() -> new SimpleDateFormat("yyyy-MM-dd\'T\'HH:mm:ss\'Z\'"));
    public static final ThreadLocal<SimpleDateFormat> dateTimeFormat
        = ThreadLocal.withInitial(() -> new SimpleDateFormat("yyyy-MM-dd HH:mm:ss"));
    public static final ThreadLocal<SimpleDateFormat> dateFormat
        = ThreadLocal.withInitial(() -> new SimpleDateFormat("yyyy-MM-dd"));
    public static final DateTimeFormatter LOCAL_DATE_TIME;

    static {
        LOCAL_DATE_TIME = new DateTimeFormatterBuilder()
            .parseCaseInsensitive()
            .append(ISO_LOCAL_DATE)
            .appendLiteral(' ')
            .append(ISO_LOCAL_TIME)
            .toFormatter();
    }

    public static String toString(Date date) {
        return dateTimeFormat.get().format(date);
    }

    public static String toTString(Date date) {
        return utcDateTimeFormat.get().format(date);
    }

    public static String toDateString(Date date) {
        return dateFormat.get().format(date);
    }

    public static String today() {
        return LocalDate.now().toString();
    }

    private DateUtil() { }
    /**
     * Timestamp to Office open xml timestamp
     *
     * @param ts the java.sql.timestamp value
     * @return Office open xml timestamp
     */
    public static double toDateTimeValue(Timestamp ts) {
        LocalDateTime ldt = ts.toLocalDateTime();
        long day = ldt.toLocalDate().toEpochDay();
        int second = ldt.toLocalTime().toSecondOfDay();
        return second / SECOND_OF_DAY + day + DAYS_1900_TO_1970;
    }

    /**
     * java.util.Date to Office open xml date
     *
     * @param date the java.util.Date value
     * @return Office open xml date
     */
    public static double toDateTimeValue(Date date) {
        return toDateTimeValue(new Timestamp(date.getTime()));
    }

    /**
     * java.util.Date to Office open xml date
     *
     * @param date the java.util.Date value
     * @return Office open xml date
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
     * Timestamp to Office open xml time-of-day
     *
     * @param ts the Timestamp value
     * @return Office open xml time-of-day
     */
    public static double toTimeValue(Timestamp ts) {
        return toTimeValue(ts.toLocalDateTime().toLocalTime());
    }

    /**
     * java.util.Date to Office open xml time-of-day
     *
     * @param date the java.util.Date value
     * @return Office open xml time-of-day
     */
    public static double toTimeValue(Date date) {
        return toTimeValue(new Timestamp(date.getTime()));
    }

    /**
     * java.sql.Time to Office open xml time-of-day
     *
     * @param time the java.sql.Time value
     * @return Office open xml time-of-day
     */
    public static double toTimeValue(Time time) {
        return toTimeValue(time.toLocalTime());
    }

    /**
     * LocalDateTime to Office open xml timestamp
     *
     * @param ldt the java.time.LocalDateTime value
     * @return Office open xml timestamp
     */
    public static double toDateTimeValue(LocalDateTime ldt) {
        long day = ldt.toLocalDate().toEpochDay();
        int second = ldt.toLocalTime().toSecondOfDay();
        return second / SECOND_OF_DAY + day + DAYS_1900_TO_1970;
    }

    /**
     * LocalDate to Office open xml date
     *
     * @param date the java.time.LocalDate value
     * @return Office open xml date
     */
    public static int toDateValue(LocalDate date) {
        return (int) date.toEpochDay() + DAYS_1900_TO_1970;
    }

    /**
     * LocalTime to Office open xml time-of-day
     *
     * @param time the LocalTime value
     * @return Office open xml time-of-day
     */
    public static double toTimeValue(LocalTime time) {
        return time.toSecondOfDay() / SECOND_OF_DAY;
    }

    /////////////////////////////number to date//////////////////////////////////

    /**
     * Office open XML timestamp to java.util.Date
     *
     * @param n the office open xml timestamp value
     * @return java.util.Date
     */
    public static java.util.Date toDate(int n) {
        if (n < DAYS_1900_TO_1970) {
            throw new UncheckedTypeException("ConstantNumber " + n + " can't convert to java.util.Date");
        }
        return Date.from(Instant.ofEpochSecond((n - DAYS_1900_TO_1970) * 86400L).minusMillis(tz));
    }

    /**
     * Office open xml timestamp to java.util.Date
     *
     * @param d the Office open xml timestamp value
     * @return java.util.Date
     */
    public static java.util.Date toDate(double d) {
        if (d - DAYS_1900_TO_1970 < .00001) {
            throw new UncheckedTypeException("ConstantNumber " + d + " can't convert to java.util.Date");
        }
        int n = (int) d, m = (int) Math.floor(((d - n) * MILLIS_OF_DAY + 0.5) / 1000);
        return Date.from(Instant.ofEpochSecond((n - DAYS_1900_TO_1970) * 86400L + m).minusMillis(tz));
    }

    public static java.util.Date toDate(String dateStr) {
        return new java.util.Date(toTimestamp(dateStr).getTime());
    }

    public static java.sql.Time toTime(double d) {
        int n = (int) d, m = (int) Math.floor(((d - n) * MILLIS_OF_DAY + 0.5) / 1000);
        return java.sql.Time.valueOf(LocalTime.ofSecondOfDay(m));
    }

    public static java.sql.Time toTime(String s) {
        LocalTime time = toLocalTime(s);
        return time != null ? java.sql.Time.valueOf(time) : null;
    }

    public static LocalTime toLocalTime(double d) {
        int n = (int) d, m = (int) Math.floor(((d - n) * MILLIS_OF_DAY + 0.5) / 1000);
        return LocalTime.ofSecondOfDay(m);
    }

    public static LocalTime toLocalTime(String s) {
        try {
            return LocalTime.parse(s);
        } catch (DateTimeParseException | NullPointerException e) {
            return null;
        }
    }

    public static java.sql.Timestamp toTimestamp(double d) {
        if (d - DAYS_1900_TO_1970 < .00001) {
            throw new UncheckedTypeException("ConstantNumber " + d + " can't convert to java.util.Date");
        }
        int n = (int) d, m = (int) Math.floor(((d - n) * MILLIS_OF_DAY + 0.5) / 1000);
        return Timestamp.from(Instant.ofEpochSecond((n - DAYS_1900_TO_1970) * 86400L + m).minusMillis(tz));
    }


    public static LocalDateTime toLocalDateTime(double d) {
        if (d - DAYS_1900_TO_1970 < .00001) {
            throw new UncheckedTypeException("ConstantNumber " + d + " can't convert to java.util.Date");
        }
        int n = (int) d, m = (int) Math.floor(((d - n) * MILLIS_OF_DAY + 0.5) / 1000);
        return LocalDateTime.ofInstant(Instant.ofEpochSecond((n - DAYS_1900_TO_1970) * 86400L + m).minusMillis(tz), ZoneId.systemDefault());
    }

    public static java.sql.Timestamp toTimestamp(int n) {
        if (n < DAYS_1900_TO_1970) {
            throw new UncheckedTypeException("ConstantNumber " + n + " can't convert to java.util.Date");
        }
        return Timestamp.from(Instant.ofEpochSecond((n - DAYS_1900_TO_1970) * 86400L).minusMillis(tz));
    }

    public static LocalDate toLocalDate(int n) {
        if (n < DAYS_1900_TO_1970) {
            throw new UncheckedTypeException("ConstantNumber " + n + " can't convert to java.util.Date");
        }
        return LocalDate.ofEpochDay(n - DAYS_1900_TO_1970);
    }

    public static java.sql.Timestamp toTimestamp(String dateStr) {
        // check format string
        if (dateStr.indexOf('/') == 4) {
            dateStr = dateStr.replace('/', '-');
        }
        return java.sql.Timestamp.valueOf(dateStr);
    }
}

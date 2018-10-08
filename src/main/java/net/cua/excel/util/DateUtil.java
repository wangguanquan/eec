package net.cua.excel.util;

import net.cua.excel.reader.UncheckedTypeException;

import java.sql.Timestamp;
import java.text.SimpleDateFormat;
import java.time.*;
import java.util.Date;
import java.util.concurrent.TimeUnit;

/**
 * For Excel
 * Created by guanquan.wang on 2017/9/21.
 */
public class DateUtil {
    static final int DAYS_1900_TO_1970 = ~(int)LocalDate.of(1900, 1, 1).toEpochDay() + 3;
    static final double SECOND_OF_DAY = (double) 24 * 60 * 60;

    public static final ThreadLocal<SimpleDateFormat> utcDateTimeFormat
            = ThreadLocal.withInitial(() -> new SimpleDateFormat("yyyy-MM-dd\'T\'HH:mm:ss\'Z\'"));
    public static final ThreadLocal<SimpleDateFormat> dateTimeFormat
            = ThreadLocal.withInitial(() -> new SimpleDateFormat("yyyy-MM-dd HH:mm:ss"));
    public static final ThreadLocal<SimpleDateFormat> dateFormat
            = ThreadLocal.withInitial(() -> new SimpleDateFormat("yyyy-MM-dd"));

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

    /**
     * days from 1900 plus current second of day
     * @param ts
     * @return
     */
    public static double toDateTimeValue(Timestamp ts) {
        LocalDateTime ldt = ts.toLocalDateTime();
        long day = ldt.toLocalDate().toEpochDay();
        int second = ldt.toLocalTime().toSecondOfDay();
        return second / SECOND_OF_DAY + day + DAYS_1900_TO_1970;
    }

    /**
     * days from 1900
     * @param date
     * @return
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

    /////////////////////////////number to date//////////////////////////////////
    // TODO time zone
    public static java.util.Date toDate(int n) {
        if (n < DAYS_1900_TO_1970) {
            throw new UncheckedTypeException("Number " + n + " can't convert to java.util.Date");
        }
        return Date.from(Instant.ofEpochSecond((n - DAYS_1900_TO_1970) * 86400).minusMillis(TimeUnit.HOURS.toMillis(8)));
    }

    public static java.util.Date toDate(double d) {
        if (d - DAYS_1900_TO_1970 < .00001) {
            throw new UncheckedTypeException("Number " + d + " can't convert to java.util.Date");
        }
        int n = (int) d, m = (int) ((d - n) * SECOND_OF_DAY);
        return Date.from(Instant.ofEpochSecond((n - DAYS_1900_TO_1970) * 86400 + m).minusMillis(TimeUnit.HOURS.toMillis(8)));
    }

    public static java.util.Date toDate(String dateStr) {
        return new java.util.Date(toTimestamp(dateStr).getTime());
    }

    public static java.sql.Time toTime(double d) {
        int n = (int) d, m = (int) ((d - n) * SECOND_OF_DAY);
        return java.sql.Time.valueOf(LocalTime.ofSecondOfDay(m));
    }

    public static java.sql.Timestamp toTimestamp(double d) {
        if (d - DAYS_1900_TO_1970 < .00001) {
            throw new UncheckedTypeException("Number " + d + " can't convert to java.util.Date");
        }
        int n = (int) d, m = (int) ((d - n) * SECOND_OF_DAY);
        return Timestamp.from(Instant.ofEpochSecond((n - DAYS_1900_TO_1970) * 86400 + m).minusMillis(TimeUnit.HOURS.toMillis(8)));
    }

    public static java.sql.Timestamp toTimestamp(int n) {
        if (n < DAYS_1900_TO_1970) {
            throw new UncheckedTypeException("Number " + n + " can't convert to java.util.Date");
        }
        return Timestamp.from(Instant.ofEpochSecond((n - DAYS_1900_TO_1970) * 86400).minusMillis(TimeUnit.HOURS.toMillis(8)));
    }

    public static java.sql.Timestamp toTimestamp(String dateStr) {
        // check format string
        if (dateStr.indexOf('/') == 4) {
            dateStr = dateStr.replace('/', '-');
        }
        return java.sql.Timestamp.valueOf(dateStr);
    }
}

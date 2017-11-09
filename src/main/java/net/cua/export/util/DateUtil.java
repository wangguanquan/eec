package net.cua.export.util;

import java.sql.Timestamp;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.util.Date;

/**
 * For Excel
 * Created by wanggq on 2017/9/21.
 */
public class DateUtil {
    static final int DAYS_1900_TO_1970 = ~(int)LocalDate.of(1900, 1, 1).toEpochDay() + 3;
    static final double SECOND_OF_DAY = (double) 24 * 60 * 60;

    public static final ThreadLocal<SimpleDateFormat> utcDateTimeFormat = new ThreadLocal<SimpleDateFormat>() {
        public SimpleDateFormat initialValue() {
            return new SimpleDateFormat("yyyy-MM-dd\'T\'HH:mm:ss\'Z\'");
        }
    };
    public static final ThreadLocal<SimpleDateFormat> dateTimeFormat = new ThreadLocal<SimpleDateFormat>() {
        public SimpleDateFormat initialValue() {
            return new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        }
    };
    public static final ThreadLocal<SimpleDateFormat> dateFormat = new ThreadLocal<SimpleDateFormat>() {
        public SimpleDateFormat initialValue() {
            return new SimpleDateFormat("yyyy-MM-dd");
        }
    };

    public static String toString(Date date) {
        return dateTimeFormat.get().format(date);
    }

    public static String toTString(Date date) {
        return utcDateTimeFormat.get().format(date);
    }

    public static String toDateString(Date date) {
        return dateFormat.get().format(date);
    }

    public static String getToday() {
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
        } else if (date instanceof java.util.Date) {
            n = (int) date.toInstant().atZone(ZoneId.systemDefault()).toLocalDate().toEpochDay() + DAYS_1900_TO_1970;
        } else {
            n = -1;
        }
        return n;
    }
}

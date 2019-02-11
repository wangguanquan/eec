package net.cua.excel.util;

import net.cua.excel.reader.UncheckedTypeException;

import java.sql.Time;
import java.sql.Timestamp;
import java.text.SimpleDateFormat;
import java.time.*;
import java.util.Date;
import java.util.TimeZone;
import java.util.concurrent.TimeUnit;

/**
 * For Excel
 * Created by guanquan.wang at 2017/9/21.
 */
public class DateUtil {
    private static final int DAYS_1900_TO_1970 = ~(int)LocalDate.of(1900, 1, 1).toEpochDay() + 3;
    private static final double SECOND_OF_DAY = (double) 24 * 60 * 60;

    // time-zone
    private static final int tz = TimeZone.getDefault().getRawOffset() / 1000 / 60 / 60;

    private static final ThreadLocal<SimpleDateFormat> utcDateTimeFormat
            = ThreadLocal.withInitial(() -> new SimpleDateFormat("yyyy-MM-dd\'T\'HH:mm:ss\'Z\'"));
    private static final ThreadLocal<SimpleDateFormat> dateTimeFormat
            = ThreadLocal.withInitial(() -> new SimpleDateFormat("yyyy-MM-dd HH:mm:ss"));
    private static final ThreadLocal<SimpleDateFormat> dateFormat
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
     * Timestamp to Office open xml timestamp
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
     * @param ts the Timestamp value
     * @return Office open xml time-of-day
     */
    public static double toTimeValue(Timestamp ts) {
        return toTimeValue(ts.toLocalDateTime().toLocalTime());
    }

    /**
     * java.util.Date to Office open xml time-of-day
     * @param date the java.util.Date value
     * @return Office open xml time-of-day
     */
    public static double toTimeValue(Date date) {
        return toTimeValue(new Timestamp(date.getTime()));
    }

    /**
     * java.sql.Time to Office open xml time-of-day
     * @param time the java.sql.Time value
     * @return Office open xml time-of-day
     */
    public static double toTimeValue(Time time) {
        return toTimeValue(time.toLocalTime());
    }

    /**
     * LocalDateTime to Office open xml timestamp
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
     * @param date the java.time.LocalDate value
     * @return Office open xml date
     */
    public static int toDateValue(LocalDate date) {
        return (int) date.toEpochDay() + DAYS_1900_TO_1970;
    }

    /**
     * LocalTime to Office open xml time-of-day
     * @param time the LocalTime value
     * @return Office open xml time-of-day
     */
    public static double toTimeValue(LocalTime time) {
        return time.toSecondOfDay() / SECOND_OF_DAY;
    }

    /////////////////////////////number to date//////////////////////////////////

    /**
     * Office open XML timestamp to java.util.Date
     * @param n the office open xml timestamp value
     * @return java.util.Date
     */
    public static java.util.Date toDate(int n) {
        if (n < DAYS_1900_TO_1970) {
            throw new UncheckedTypeException("ConstantNumber " + n + " can't convert to java.util.Date");
        }
        return Date.from(Instant.ofEpochSecond((n - DAYS_1900_TO_1970) * 86400L).minusMillis(TimeUnit.HOURS.toMillis(tz)));
    }

    /**
     * Office open xml timestamp to java.util.Date
     * @param d the Office open xml timestamp value
     * @return java.util.Date
     */
    public static java.util.Date toDate(double d) {
        if (d - DAYS_1900_TO_1970 < .00001) {
            throw new UncheckedTypeException("ConstantNumber " + d + " can't convert to java.util.Date");
        }
        int n = (int) d, m = (int) ((d - n) * SECOND_OF_DAY);
        return Date.from(Instant.ofEpochSecond((n - DAYS_1900_TO_1970) * 86400L + m).minusMillis(TimeUnit.HOURS.toMillis(tz)));
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
            throw new UncheckedTypeException("ConstantNumber " + d + " can't convert to java.util.Date");
        }
        int n = (int) d, m = (int) ((d - n) * SECOND_OF_DAY);
        return Timestamp.from(Instant.ofEpochSecond((n - DAYS_1900_TO_1970) * 86400L + m).minusMillis(TimeUnit.HOURS.toMillis(tz)));
    }

    public static java.sql.Timestamp toTimestamp(int n) {
        if (n < DAYS_1900_TO_1970) {
            throw new UncheckedTypeException("ConstantNumber " + n + " can't convert to java.util.Date");
        }
        return Timestamp.from(Instant.ofEpochSecond((n - DAYS_1900_TO_1970) * 86400L).minusMillis(TimeUnit.HOURS.toMillis(tz)));
    }

    public static java.sql.Timestamp toTimestamp(String dateStr) {
        // check format string
        if (dateStr.indexOf('/') == 4) {
            dateStr = dateStr.replace('/', '-');
        }
        return java.sql.Timestamp.valueOf(dateStr);
    }

    /**
     * Check leap year
     * a year divisible by 4 is a leap year;
     * with the exception that a year divisible by 100 is not a leap year (e.g. 1900 was no leap year);
     * with the exception that a year divisible by 400 is a leap year (e.g. 2000 was a leap year).
     * @param year the year
     * @return true if leap year
     */
    public static boolean isLeapYear(int year) {
        return year % 4 == 0 && year % 100 != 0 || year % 400 == 0 && year % 3200 != 0;
    }

    // --- BIFF Timestamp support

    /**
     * The time stamp field is an unsigned 64-bit integer value that contains the time elapsed
     * since 1601-Jan-01 00:00:00 (Gregorian calendar3).
     * One unit of this value is equal to 100 nanoseconds (10–7 seconds).
     * That means, each second the time stamp value will be increased by 10 million units.
     */
    public static final class Biff {
        // One unit of this value is equal to 100 nanoseconds (10–7 seconds).
        private static final int base_nano = 10_000_000;
        private static final int base_year = 1601;
        // Days of month
        private static final int[] month_table = {
            31, 28, 31,
            30, 31, 30,
            31, 31, 30,
            31, 30, 31,
        };

        /**
         * biff timestamp to LocalDateTime
         * @param time the biff timestamp value
         * @return java.time.LocalDateTime
         */
        public static LocalDateTime toLocalDateTime(long time) {
            // Fractional amount of a second
            int frac = (int) (time % base_nano);
            // Remaining entire seconds
            long t1 = time / base_nano;

            // Seconds in a minute
            int sec = (int) (t1 % 60);
            // Remaining entire minutes
            long t2 = t1 / 60;

            // Minutes in an hour
            int min = (int) (t2 % 60);
            // Remaining entire hours
            int t3 = (int) (t2 / 60);

            // Hours in a day
            int hour = t3 % 24;
            // Remaining entire days
            int t4 = t3 / 24;

            // Entire years from 1601-Jan-01
            int year = base_year;
            for ( ; (t4 -= isLeapYear(year++) ? 366 : 365) >= 365; );

            int k = isLeapYear(year) ? 1 : 0;

            // Entire months from 1977-Jan-01
            int t5 = t4
                // number of full months in t5
                , m = 0;
            // number of days from 1977-Jan-01 to 1977-Apr-01
            for ( ; (t5 -= m == 1 ? month_table[m++] + k : month_table[m++]) >= 28; );

            int month = 1 + m;

            // Resulting day of month April
            int day = 1 + t5;

            return LocalDateTime.of(year, month, day, hour, min, sec, frac * 100);
        }

        /**
         * biff timestamp to java.sql.Timestamp
         * @param time the biff timestamp value
         * @return java.sql.Timestamp
         */
        public static java.sql.Timestamp toTimestamp(long time) {
            return Timestamp.from(toLocalDateTime(time).toInstant(ZoneOffset.ofHours(tz)));
        }

        /**
         * biff timestamp to LocalDate
         * @param time the biff timestamp value
         * @return java.time.LocalDate
         */
        public static LocalDate toLocalDate(long time) {
            int t4 = (int) (time / 864_000_000_000L);
            // Entire years from 1601-Jan-01
            int year = base_year;
            // Remaining days in year 1977
            // number of full years in t4
            for ( ; (t4 -= isLeapYear(year++) ? 366 : 365) >= 365; );

            int k = isLeapYear(year) ? 1 : 0;

            // Entire months from 1977-Jan-01
            int t5 = t4
                // number of full months in t5
                , m = 0;
            for ( ; (t5 -= m == 1 ? month_table[m++] + k : month_table[m++]) >= 28; );

            int month = 1 + m;

            // Resulting day of month April
            int day = 1 + t5;

            return LocalDate.of(year, month, day);
        }

        /**
         * biff timestamp to LocalTime
         * @param time the biff timestamp value
         * @return java.time.LocalTime
         */
        public static LocalTime toLocalTime(long time) {
            // Fractional amount of a second
            int frac = (int) (time % base_nano);
            // Remaining entire seconds
            long t1 = time / base_nano;

            // Seconds in a minute
            int sec = (int) (t1 % 60);
            // Remaining entire minutes
            long t2 = t1 / 60;

            // Minutes in an hour
            int min = (int) (t2 % 60);
            // Remaining entire hours
            int t3 = (int) (t2 / 60);

            // Hours in a day
            int hour = t3 % 24;

            return LocalTime.of(hour, min, sec, frac);
        }

        /**
         * biff timestamp value to java.sql.Time
         * @param time the biff timestamp value
         * @return java.sql.Time
         */
        public static java.sql.Time toTime(long time) {
            return Time.valueOf(toLocalTime(time));
        }

        /**
         * biff timestamp value to java.util.Date
         * @param time the biff timestamp value
         * @return java.util.Date
         */
        public static java.util.Date toDate(long time) {
            return toTimestamp(time);
        }

        /**
         * Timestamp to unsigned 64-bit integer value
         * @param ts the Timestamp
         * @return unsigned 64-bit value
         */
        public static long toDateTimeValue(Timestamp ts) {
            return toDateTimeValue(ts.toLocalDateTime());
        }

        /**
         * java.util.Date to unsigned 64-bit integer value
         * @param date the java.util.Date
         * @return unsigned 64-bit value
         */
        public static long toDateTimeValue(Date date) {
            return toDateTimeValue(new Timestamp(date.getTime()));
        }

        /**
         * LocalDateTime to unsigned 64-bit integer value
         * @param ldt the local-date-time
         * @return unsigned 64-bit value
         */
        public static long toDateTimeValue(LocalDateTime ldt) {
            // Days of month
            int day = ldt.getDayOfMonth() - 1;
            // Month of year from 1 to 12.
            int month = ldt.getMonthValue() - 1;
            int year = ldt.getYear();
            long t0 = day;
            int k = isLeapYear(year) ? 1 : 0;
            for ( ; month-- > 0; t0 += month == 1 ? month_table[month] + k : month_table[month]);
            // number of full years from 1601-Jan-01
            for ( ; year-- > base_year; t0 += isLeapYear(year) ? 366 : 365);

            return  (((t0 * 24 + ldt.getHour()) * 60 + ldt.getMinute()) * 60 + ldt.getSecond()) * base_nano + ldt.getNano() / 100;
        }
    }

}

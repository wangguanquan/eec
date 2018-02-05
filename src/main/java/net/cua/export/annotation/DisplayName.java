package net.cua.export.annotation;

import java.lang.annotation.*;

/**
 * Created by wanggq at 2018-01-30 13:23
 */
@Target({ ElementType.FIELD })
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface DisplayName {
    /**
     * header cell
     * @return
     */
    String value() default "";

    /**
     * share body string
     * @return
     */
    boolean share() default false;
}

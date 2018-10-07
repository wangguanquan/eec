package net.cua.excel.annotation;

import java.lang.annotation.*;

/**
 * Created by guanquan.wang at 2018-01-30 15:09
 */
@Target({ ElementType.FIELD })
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface NotExport {
    /**
     * describe why not export
     * @return
     */
    String value() default "";
}

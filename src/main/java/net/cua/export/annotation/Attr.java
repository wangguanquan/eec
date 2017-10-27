package net.cua.export.annotation;

import java.lang.annotation.*;

/**
 * Created by wanggq on 2017/9/21.
 */
@Target({ ElementType.FIELD })
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface Attr {
    String[] name();
    String[] value() default {};
    NS namespace() default @NS("-");
}

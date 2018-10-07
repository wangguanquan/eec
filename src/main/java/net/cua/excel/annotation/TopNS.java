package net.cua.excel.annotation;

import java.lang.annotation.*;

/**
 * Created by guanquan.wang on 2017/9/21.
 */
@Target(ElementType.TYPE)
@Retention(RetentionPolicy.RUNTIME)
@Inherited
@Documented
public @interface TopNS {
    String[] prefix();
    String[] uri() default {};
    String value();
}
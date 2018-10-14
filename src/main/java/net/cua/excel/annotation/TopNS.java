package net.cua.excel.annotation;

import java.lang.annotation.*;

/**
 * Top namespace
 * Created by guanquan.wang at 2017/9/21.
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
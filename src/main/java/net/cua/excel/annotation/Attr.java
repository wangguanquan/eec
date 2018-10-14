package net.cua.excel.annotation;

import java.lang.annotation.*;

/**
 * Xml Attribute
 * Created by guanquan.wang at 2017/9/21.
 */
@Target({ ElementType.FIELD })
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface Attr {
    /**
     * attribute name
     */
    String[] name();

    /**
     * attribute value
     * @return
     */
    String[] value() default {};

    /**
     * namespace
     * @return
     */
    NS namespace() default @NS("-");
}

package net.cua.excel.annotation;

import java.lang.annotation.*;

/**
 * 指定字段不做导入映射
 * Created by guanquan.wang at 2018-10-24 09:29
 */
@Target({ ElementType.FIELD })
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface NotImport {
    /**
     * 添加不导入的原因，方便开发者阅读
     * @return 原因
     */
    String value() default "";
}

package org.ttzero.excel.annotation;

import org.ttzero.excel.processor.StyleProcessor;

import java.lang.annotation.ElementType;
import java.lang.annotation.Inherited;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * customer style of specified row
 *
 * @author suyl at 2022-03-23 17:38
 *
 */
@Target(ElementType.TYPE)
@Retention(RetentionPolicy.RUNTIME)
@Inherited
public @interface StyleDesign {
    Class<? extends StyleProcessor> using() default StyleProcessor.None.class;
}


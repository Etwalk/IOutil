package annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.METHOD, ElementType.FIELD})
/**
 * index值从0开始，如果超过了excel每行的单元格数，则结果为null
 */
public @interface ParaseExcelAnnotation {
    int index() default 0;
}

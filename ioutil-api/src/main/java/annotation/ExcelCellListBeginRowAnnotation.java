package annotation;

import java.lang.annotation.*;

/**
 * @author Sherlock.Wu
 * @date 2019/11/25
 */
@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.TYPE})
public @interface ExcelCellListBeginRowAnnotation{
    /**
     * 结束的标志行是为空的时候
     * list开始的标志
     * 和row()互相排斥
     * 该数据从那一行开始，到row cell 为空的时候结束
     * @return
     */
    int beginRow();
}

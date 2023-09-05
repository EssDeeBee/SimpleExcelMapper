package hckrn.simpleexcelmapper.annotation;


import hckrn.simpleexcelmapper.format.ExcelColumnDataFormat;
import org.apache.poi.ss.usermodel.IndexedColors;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Target({ElementType.TYPE})
@Retention(RetentionPolicy.RUNTIME)
public @interface DocumentExcel {

    String name() default "";

    int cellPosition() default -1;

    ExcelColumnDataFormat cellTypePattern() default ExcelColumnDataFormat.NONE;

    IndexedColors cellColor() default IndexedColors.AUTOMATIC;
}

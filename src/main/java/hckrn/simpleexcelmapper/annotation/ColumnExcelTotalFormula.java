package hckrn.simpleexcelmapper.annotation;


import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Target({ElementType.METHOD})
@Retention(RetentionPolicy.RUNTIME)
public @interface ColumnExcelTotalFormula {
    boolean useValue() default false;
    int position();
    ColumnExcelStyle cellStyle() default @ColumnExcelStyle;
}

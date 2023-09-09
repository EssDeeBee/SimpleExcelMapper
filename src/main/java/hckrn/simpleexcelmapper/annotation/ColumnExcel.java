package hckrn.simpleexcelmapper.annotation;

import hckrn.simpleexcelmapper.format.ExcelColumnCellTextColor;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Target({ElementType.FIELD})
@Retention(RetentionPolicy.RUNTIME)
public @interface ColumnExcel {
    String[] applyNames() default {};

    int position();

    ColumnExcelStyle headerStyle() default @ColumnExcelStyle(fontColor = ExcelColumnCellTextColor.BLACK,
            isCentreAlignment = true,
            isFontBold = true,
            fontSize = 14,
            isWrapText = true);

    ColumnExcelStyle cellStyle() default @ColumnExcelStyle;
}

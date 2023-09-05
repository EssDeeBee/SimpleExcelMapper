package hckrn.simpleexcelmapper.annotation;

import hckrn.simpleexcelmapper.format.ExcelColumnCellTextColor;

import java.lang.annotation.*;

@Target({ElementType.METHOD})
@Retention(RetentionPolicy.RUNTIME)
@Inherited
public @interface ColumnExcelFormula {
    String name() default "";

    int position();

    ColumnExcelStyle headerStyle() default @ColumnExcelStyle(fontColor = ExcelColumnCellTextColor.BLACK,
            isCentreAlignment = true,
            isFontBold = true,
            fontSize = 14,
            isWrapText = true);

    ColumnExcelStyle cellStyle() default @ColumnExcelStyle(fontSize = 8);
}

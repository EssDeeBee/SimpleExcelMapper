package hckrn.simpleexcelmapper.annotation;


import hckrn.simpleexcelmapper.format.ExcelColumnCellTextColor;
import hckrn.simpleexcelmapper.format.ExcelColumnDataFormat;
import hckrn.simpleexcelmapper.format.ExcelColumnFont;
import org.apache.poi.ss.usermodel.IndexedColors;

import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;

@Retention(RetentionPolicy.RUNTIME)
public @interface ColumnExcelStyle {
    ExcelColumnDataFormat cellTypePattern() default ExcelColumnDataFormat.NONE;
    IndexedColors cellColor() default IndexedColors.AUTOMATIC;
    boolean isWrapText() default false;
    boolean isCentreAlignment() default false;
    boolean isFramed() default true;
    ExcelColumnFont fontName() default ExcelColumnFont.DEFAULT;
    short fontSize() default -1;
    boolean isFontBold() default false;
    ExcelColumnCellTextColor fontColor() default ExcelColumnCellTextColor.AUTOMATIC;
}

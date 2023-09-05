package hckrn.simpleexcelmapper.format;

import lombok.Getter;

@Getter
public enum ExcelColumnDataFormat {

    RUR("#,##0.00 [$р.-419];-#,##0.00 [$р.-419]"),
    PERCENTAGE("0.00%"),
    NUMBER("#,##0.00"),
    DATE("D MMM YY;@"),
    TIME("[$-F400]H:MM:SS AM/PM"),
    NONE("");

    private final String formatPattern;

    ExcelColumnDataFormat(String formatPattern) {
        this.formatPattern = formatPattern;
    }
}

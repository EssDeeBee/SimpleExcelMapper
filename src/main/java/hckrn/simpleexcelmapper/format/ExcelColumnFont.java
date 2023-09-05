package hckrn.simpleexcelmapper.format;

import lombok.Getter;

@Getter
public enum ExcelColumnFont {
    DEFAULT(""),
    CALIBRI_LIGHT("Calibri Light"),
    TAHOMA("Tahoma");


    private final String fontName;

    ExcelColumnFont(String fontName) {
        this.fontName = fontName;
    }
}

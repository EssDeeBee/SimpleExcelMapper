package hckrn.simpleexcelmapper.format;

import lombok.Getter;

@Getter
public enum ExcelColumnCellTextColor {
    WHITE((short) 0x09),
    AUTOMATIC((short) 64),
    BLACK((short) 8);

    private final short colorIndex;

    ExcelColumnCellTextColor(short colorIndex) {
        this.colorIndex = colorIndex;
    }
}

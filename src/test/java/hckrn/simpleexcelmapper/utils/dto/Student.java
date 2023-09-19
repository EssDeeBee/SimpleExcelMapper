package hckrn.simpleexcelmapper.utils.dto;

import hckrn.simpleexcelmapper.annotation.*;
import hckrn.simpleexcelmapper.format.ExcelColumnCellTextColor;
import hckrn.simpleexcelmapper.format.ExcelColumnDataFormat;
import lombok.Data;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.util.CellAddress;

import java.time.LocalDate;


@Data
@DocumentExcel
public class Student {
    @ColumnExcel(applyNames = "Student Id", position = 0,
            headerStyle = @ColumnExcelStyle(fontSize = 15, cellColor = IndexedColors.CORAL),
            cellStyle = @ColumnExcelStyle(fontColor = ExcelColumnCellTextColor.WHITE, cellColor = IndexedColors.BLUE))
    private Integer studentId;

    @ColumnExcel(applyNames = "Name", position = 1,
            headerStyle = @ColumnExcelStyle(fontSize = 13, isCentreAlignment = false, cellColor = IndexedColors.BROWN))
    private String name;

    @ColumnExcel(applyNames = "Age", position = 2)
    private Integer age;

    @ColumnExcel(applyNames = "Admission Date ", position = 3,
            cellStyle = @ColumnExcelStyle(cellTypePattern = ExcelColumnDataFormat.DATE))
    private LocalDate admissionDate;

    @ColumnExcelFormula(name = "Age n Id", position = 4)
    public String idAndAge(int rowNum) {
        return "IFERROR(" + new CellAddress(rowNum, 0).formatAsString()
                + "*"
                + new CellAddress(rowNum, 2).formatAsString() +
                ",0) "
                ;
    }

    @ColumnExcelTotalFormula(position = 0)
    public  static String studentIdsSum(int firstRowNum, int lastRowNum) {
        return "IFERROR(SUM("
                + new CellAddress(firstRowNum, 0).formatAsString()
                + ":"
                + new CellAddress(lastRowNum - 1, 0).formatAsString()
                +
                "),0)";
    }
}

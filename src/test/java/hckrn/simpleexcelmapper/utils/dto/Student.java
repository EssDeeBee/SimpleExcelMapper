package hckrn.simpleexcelmapper.utils.dto;

import hckrn.simpleexcelmapper.annotation.ColumnExcel;
import hckrn.simpleexcelmapper.annotation.ColumnExcelFormula;
import hckrn.simpleexcelmapper.annotation.ColumnExcelStyle;
import hckrn.simpleexcelmapper.annotation.ColumnExcelTotalFormula;
import lombok.Data;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.util.CellAddress;


@Data
public class Student {
    @ColumnExcel(applyNames = "Student ID", position = 0,
            headerStyle = @ColumnExcelStyle(fontSize = 40, cellColor = IndexedColors.CORAL))
    private Integer studentId;

    @ColumnExcel(applyNames = "Name", position = 1)
    private String name;

    @ColumnExcel(applyNames = "Age", position = 2)
    private Integer age;

    @ColumnExcelFormula(name = "Age n Id", position = 3)
    public String idAndAge(int rowNum) {
        return "IFERROR(" + new CellAddress(rowNum, 0).formatAsString()
                + "*"
                + new CellAddress(rowNum, 2).formatAsString() +
                ",0) "
                ;
    }

    @ColumnExcelTotalFormula(position = 0)
    public static String studentIdsSum(int firstRowNum, int lastRowNum) {
        return "IFERROR(SUM("
                + new CellAddress(firstRowNum, 0).formatAsString()
                + ":"
                + new CellAddress(lastRowNum - 1, 0).formatAsString()
                +
                "),0)";
    }
}

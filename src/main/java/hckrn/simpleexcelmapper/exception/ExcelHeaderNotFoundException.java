package hckrn.simpleexcelmapper.exception;

import lombok.Getter;
import org.apache.poi.ss.usermodel.Row;

@Getter
public class ExcelHeaderNotFoundException extends RuntimeException {

    private String sheetName;
    private Integer excelRowNumber;
    private String headerName;
    private String message;

    public ExcelHeaderNotFoundException(Row excelRow, String headerName) {
        if (excelRow != null) {
            this.sheetName = excelRow.getSheet().getSheetName();
            this.excelRowNumber = excelRow.getRowNum();

        }
        this.headerName = headerName;
        this.message = "Header not found. "
                + "Sheet: " + sheetName
                + " Row: " + excelRowNumber
                + " Header: " + headerName;
    }

    public ExcelHeaderNotFoundException(Row excelRow, String headerName, String additionalInfo) {
        this(excelRow, headerName);
        this.message += "\n" + additionalInfo;
    }
}

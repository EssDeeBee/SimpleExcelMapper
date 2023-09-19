package hckrn.simpleexcelmapper.exception;

import lombok.Getter;
import org.apache.poi.ss.usermodel.Row;

import java.util.Optional;

@Getter
public class ExcelHeaderNotFoundException extends RuntimeException {

    private final String sheetName;
    private final Integer excelRowNumber;
    private final String headerName;
    private final String message;

    public ExcelHeaderNotFoundException(Row excelRow, String headerName) {
        this.sheetName = Optional.ofNullable(excelRow).map(row -> row.getSheet().getSheetName()).orElse(null);
        this.excelRowNumber = Optional.ofNullable(excelRow).map(Row::getRowNum).orElse(null);

        this.headerName = headerName;
        this.message = "Header not found. "
                + "Sheet: " + sheetName
                + " Row: " + excelRowNumber
                + " Header: " + headerName;
    }
}

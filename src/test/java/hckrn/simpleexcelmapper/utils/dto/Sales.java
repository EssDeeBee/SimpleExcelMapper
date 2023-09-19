package hckrn.simpleexcelmapper.utils.dto;

import hckrn.simpleexcelmapper.annotation.*;
import lombok.Data;
import lombok.experimental.Accessors;
import org.apache.poi.ss.util.CellAddress;

import java.time.LocalDate;

import static hckrn.simpleexcelmapper.format.ExcelColumnCellTextColor.WHITE;
import static hckrn.simpleexcelmapper.format.ExcelColumnDataFormat.*;
import static org.apache.poi.ss.usermodel.IndexedColors.*;

@Data
@Accessors(chain = true)
public class Sales {

    @ColumnExcel(position = 0, applyNames = {"Date"},
            headerStyle = @ColumnExcelStyle(fontColor = WHITE, cellColor = DARK_BLUE, isCentreAlignment = true),
            cellStyle = @ColumnExcelStyle(cellColor = GREY_25_PERCENT, cellTypePattern = DATE))
    private LocalDate date;

    @ColumnExcel(position = 1, applyNames = {"Sold"},
            headerStyle = @ColumnExcelStyle(fontColor = WHITE, cellColor = DARK_BLUE, isCentreAlignment = true),
            cellStyle = @ColumnExcelStyle(cellColor = GREY_25_PERCENT))
    private Integer sold;

    @ColumnExcel(position = 2, applyNames = {"Price Per Unit (USD)"},
            headerStyle = @ColumnExcelStyle(fontColor = WHITE, cellColor = DARK_BLUE, isCentreAlignment = true),
            cellStyle = @ColumnExcelStyle(cellColor = GREY_25_PERCENT, cellTypePattern = USD))
    private Double pricePerUnit;

    @ColumnExcelFormula(position = 3, name = "Total Sales (USD)",
            headerStyle = @ColumnExcelStyle(fontColor = WHITE, cellColor = DARK_BLUE, isCentreAlignment = true),
            cellStyle = @ColumnExcelStyle(cellColor = GREY_25_PERCENT, cellTypePattern = USD))
    public String sales(int rowNum) {
        return new CellAddress(rowNum, 1).formatAsString()
                + "*"
                + new CellAddress(rowNum, 2).formatAsString();
    }

    @ColumnExcelTotalFormula(position = 0,
            cellStyle = @ColumnExcelStyle(cellColor = LIGHT_BLUE))
    public static String total(int firstRowNum, int lastRowNum) {
        return "CONCATENATE(\"Total\")";
    }

    @ColumnExcelTotalFormula(position = 1,
            cellStyle = @ColumnExcelStyle(cellColor = LIGHT_BLUE))
    public static String unitsSold(int firstRowNum, int lastRowNum) {
        return "SUM(" + new CellAddress(firstRowNum, 1).formatAsString() + ":"
                + new CellAddress(lastRowNum - 1, 1).formatAsString() + ")";
    }

    @ColumnExcelTotalFormula(position = 2,
            cellStyle = @ColumnExcelStyle(cellColor = LIGHT_BLUE))
    public static String empty(int firstRowNum, int lastRowNum) {
        return "CONCATENATE(\"\")";
    }

    @ColumnExcelTotalFormula(position = 3,
            cellStyle = @ColumnExcelStyle(isCentreAlignment = false, cellColor = LIGHT_BLUE, cellTypePattern = USD))
    public static String totalSales(int firstRowNum, int lastRowNum) {
        return "SUM(" + new CellAddress(firstRowNum, 3).formatAsString() + ":"
                + new CellAddress(lastRowNum - 1, 3).formatAsString() + ")";
    }
}

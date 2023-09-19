package hckrn.simpleexcelmapper.utils.dto;


import hckrn.simpleexcelmapper.annotation.ColumnExcel;
import hckrn.simpleexcelmapper.annotation.ColumnExcelFormula;
import hckrn.simpleexcelmapper.annotation.ColumnExcelStyle;
import hckrn.simpleexcelmapper.annotation.DocumentExcel;
import hckrn.simpleexcelmapper.format.ExcelColumnDataFormat;
import lombok.Data;
import lombok.experimental.Accessors;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.util.CellAddress;

import java.time.LocalDate;

@Data
@DocumentExcel
@Accessors(chain = true)
public class Demo {
    @ColumnExcel(position = 0, applyNames = {"Ref"}, headerStyle = @ColumnExcelStyle(cellColor = IndexedColors.LIGHT_GREEN, isCentreAlignment = true, isWrapText = true))
    private Integer ref;
    @ColumnExcel(position = 1, applyNames = {"Date"},
            headerStyle = @ColumnExcelStyle(cellColor = IndexedColors.LIGHT_GREEN, isCentreAlignment = true, isWrapText = true),
            cellStyle = @ColumnExcelStyle(cellTypePattern = ExcelColumnDataFormat.DATE))
    private LocalDate date;
    @ColumnExcel(position = 2, applyNames = {"Home"}, headerStyle = @ColumnExcelStyle(cellColor = IndexedColors.LIGHT_GREEN, isCentreAlignment = true, isWrapText = true))
    private String home;
    @ColumnExcel(position = 3, applyNames = {"Away"}, headerStyle = @ColumnExcelStyle(cellColor = IndexedColors.LIGHT_GREEN, isCentreAlignment = true, isWrapText = true))
    private String away;
    @ColumnExcel(position = 4, applyNames = {"League"}, headerStyle = @ColumnExcelStyle(cellColor = IndexedColors.LIGHT_GREEN, isCentreAlignment = true, isWrapText = true))
    private String league;
    @ColumnExcel(position = 5, applyNames = {"Last 5: FT - Over 0.5"}, headerStyle = @ColumnExcelStyle(cellColor = IndexedColors.PALE_BLUE, isCentreAlignment = true, isWrapText = true),
            cellStyle = @ColumnExcelStyle(cellTypePattern = ExcelColumnDataFormat.PERCENTAGE))
    private Double ftOver05;
    @ColumnExcel(position = 6, applyNames = {"Last 5: FT - Over 1.5"}, headerStyle = @ColumnExcelStyle(cellColor = IndexedColors.PALE_BLUE, isCentreAlignment = true, isWrapText = true),
            cellStyle = @ColumnExcelStyle(cellTypePattern = ExcelColumnDataFormat.PERCENTAGE))
    private Double ftOver15;
    @ColumnExcel(position = 7, applyNames = {"Last 5: FT - Over 2.5"}, headerStyle = @ColumnExcelStyle(cellColor = IndexedColors.PALE_BLUE, isCentreAlignment = true, isWrapText = true),
            cellStyle = @ColumnExcelStyle(cellTypePattern = ExcelColumnDataFormat.PERCENTAGE))
    private Double ftOver25;
    @ColumnExcel(position = 8, applyNames = {"Last 5: FT - BTTS"}, headerStyle = @ColumnExcelStyle(cellColor = IndexedColors.PALE_BLUE, isCentreAlignment = true, isWrapText = true),
            cellStyle = @ColumnExcelStyle(cellTypePattern = ExcelColumnDataFormat.PERCENTAGE))
    private Double ftOver35;
    @ColumnExcel(position = 9, applyNames = {"Last 5: Home - Ave goals scored"}, headerStyle = @ColumnExcelStyle(cellColor = IndexedColors.PALE_BLUE, isCentreAlignment = true, isWrapText = true),
            cellStyle = @ColumnExcelStyle(cellTypePattern = ExcelColumnDataFormat.NUMBER))
    private Double ftBtt;
    @ColumnExcel(position = 10, applyNames = {"Last 5: Home - Ave goals scored"}, headerStyle = @ColumnExcelStyle(cellColor = IndexedColors.PALE_BLUE, isCentreAlignment = true, isWrapText = true),
            cellStyle = @ColumnExcelStyle(cellTypePattern = ExcelColumnDataFormat.NUMBER))
    private Double homeGoalScores;
    @ColumnExcel(position = 11, applyNames = {"Last 5: Home - Ave goals Cc"}, headerStyle = @ColumnExcelStyle(cellColor = IndexedColors.PALE_BLUE, isCentreAlignment = true, isWrapText = true),
            cellStyle = @ColumnExcelStyle(cellTypePattern = ExcelColumnDataFormat.NUMBER))
    private Double homeGoalsCc;
    @ColumnExcel(position = 12, applyNames = {"Last 5: Away - Ave goals scored"}, headerStyle = @ColumnExcelStyle(cellColor = IndexedColors.PALE_BLUE, isCentreAlignment = true, isWrapText = true),
            cellStyle = @ColumnExcelStyle(cellTypePattern = ExcelColumnDataFormat.NUMBER))
    private Double awayGoalScores;
    @ColumnExcel(position = 13, applyNames = {"Last 5: Away - Ave goals Cc"}, headerStyle = @ColumnExcelStyle(cellColor = IndexedColors.PALE_BLUE, isCentreAlignment = true, isWrapText = true),
            cellStyle = @ColumnExcelStyle(cellTypePattern = ExcelColumnDataFormat.NUMBER))
    private Double awayGoalsCc;
    @ColumnExcel(position = 14, applyNames = {"Last 5: Home - Goals scored"}, headerStyle = @ColumnExcelStyle(cellColor = IndexedColors.PALE_BLUE, isCentreAlignment = true, isWrapText = true),
            cellStyle = @ColumnExcelStyle(cellTypePattern = ExcelColumnDataFormat.PERCENTAGE))
    private Double homeGoalsScored;

    @ColumnExcelFormula(position = 15, name = "Avg goal score", headerStyle = @ColumnExcelStyle(cellColor = IndexedColors.PALE_BLUE, isCentreAlignment = true, isWrapText = true),
            cellStyle = @ColumnExcelStyle(cellTypePattern = ExcelColumnDataFormat.NUMBER))
    public String avgScores(int row) {
        return "AVERAGE(" + new CellAddress(row, 10).formatAsString()
                + ":"
                + new CellAddress(row, 13).formatAsString() + ")";
    }

}

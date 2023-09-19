package hckrn.simpleexcelmapper.utils;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.List;
import java.util.Map;

public interface ExcelMapper {
    <T> Map<String, List<T>> mapWorkbookToObjs(Workbook workbook, Class<T> tObject);

    <T> List<T> mapSheetToObjs(Sheet sheet, Class<T> tObject, boolean ignoreFirstRow);

    <T> T mapRowToObj(Row row, Class<T> tClass);

    <T> Workbook createWorkbookFromObject(List<T> reportObjects);

    <T> Workbook createReportWorkbook(List<T> reportObjects, int startRowNumber);

    <T> Workbook createWorkbookFromObject(List<T> reportObjects, int startRowNumber, String sheetName);
}

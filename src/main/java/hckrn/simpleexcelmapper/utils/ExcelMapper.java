package hckrn.simpleexcelmapper.utils;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.List;
import java.util.Map;

public interface ExcelMapper {
    /**
     * @param workbook
     * @param tObject
     * @param <T>
     * @return
     */
    <T> Map<String, List<T>> mapWorkbookToObjs(Workbook workbook, Class<T> tObject);

    /**
     * @param sheet
     * @param tObject
     * @param ignoreFirstRow
     * @param <T>
     * @return
     */
    <T> List<T> mapSheetToObjs(Sheet sheet, Class<T> tObject, boolean ignoreFirstRow);

    /**
     * @param row
     * @param tClass
     * @param <T>
     * @return
     */
    <T> T mapRowToObj(Row row, Class<T> tClass);

    /**
     * @param reportObjects
     * @param <T>
     * @return
     */
    <T> Workbook createWorkbookFromObject(List<T> reportObjects);

    /**
     * @param reportObjects
     * @param startRowNumber
     * @param <T>
     * @return
     */
    <T> Workbook createReportWorkbook(List<T> reportObjects, int startRowNumber);

    /**
     * @param reportObjects
     * @param startRowNumber
     * @param sheetName
     * @param <T>
     * @return
     */
    <T> Workbook createWorkbookFromObject(List<T> reportObjects, int startRowNumber, String sheetName);
}

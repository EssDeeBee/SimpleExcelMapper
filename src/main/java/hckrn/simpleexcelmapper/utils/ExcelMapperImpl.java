package hckrn.simpleexcelmapper.utils;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.List;
import java.util.Map;

public class ExcelMapperImpl implements ExcelMapper {
    private final ExcelToObjMapper excelToObjMapper = new ExcelToObjMapper();
    private final ObjToExcelMapper objToExcelMapper = new ObjToExcelMapper();


    @Override
    public <T> Map<String, List<T>> mapWorkbookToObjs(Workbook workbook, Class<T> tObject) {
        return excelToObjMapper.mapWorkbookToObjs(workbook, tObject);
    }

    @Override
    public <T> List<T> mapSheetToObjs(Sheet sheet, Class<T> tObject, boolean ignoreFirstRow) {
        return excelToObjMapper.mapSheetToObjs(sheet, tObject, ignoreFirstRow);
    }

    @Override
    public <T> T mapRowToObj(Row row, Class<T> tClass) {
        return excelToObjMapper.mapRowToObj(row, tClass);
    }

    @Override
    public <T> Workbook createWorkbookFromObject(List<T> reportObjects) {
        return objToExcelMapper.createWorkbookFromObject(reportObjects);
    }

    @Override
    public <T> Workbook createReportWorkbook(List<T> reportObjects, int startRowNumber) {
        return objToExcelMapper.createReportWorkbook(reportObjects, startRowNumber);
    }

    @Override
    public <T> Workbook createWorkbookFromObject(List<T> reportObjects, int startRowNumber, String sheetName) {
        return objToExcelMapper.createWorkbookFromObject(reportObjects, startRowNumber, sheetName);
    }
}

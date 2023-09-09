package hckrn.simpleexcelmapper.utils;


import hckrn.simpleexcelmapper.annotation.ColumnExcel;
import hckrn.simpleexcelmapper.annotation.ColumnExcelFormula;
import hckrn.simpleexcelmapper.annotation.ColumnExcelStyle;
import hckrn.simpleexcelmapper.annotation.ColumnExcelTotalFormula;
import hckrn.simpleexcelmapper.exception.ExcelMapperException;
import hckrn.simpleexcelmapper.format.ExcelColumnCellTextColor;
import hckrn.simpleexcelmapper.format.ExcelColumnDataFormat;
import hckrn.simpleexcelmapper.format.ExcelColumnFont;
import lombok.SneakyThrows;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.beans.IntrospectionException;
import java.beans.Introspector;
import java.beans.PropertyDescriptor;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.lang.reflect.Modifier;
import java.math.BigDecimal;
import java.sql.Date;
import java.time.LocalDate;
import java.time.ZoneId;
import java.util.List;

import static java.util.Objects.nonNull;

@Slf4j
class ObjToExcelMapper {

    <T> Workbook createWorkbookFromObject(List<T> reportObjects) {
        return createWorkbookFromObject(reportObjects, 0, "Report_" + LocalDate.now());
    }

    <T> Workbook createReportWorkbook(List<T> reportObjects, int startRowNumber) {
        return createWorkbookFromObject(reportObjects, startRowNumber, "Report_" + LocalDate.now());
    }

    <T> Workbook createWorkbookFromObject(List<T> reportObjects, int startRowNumber, String sheetName) {

        if (nonNull(reportObjects) && !reportObjects.isEmpty()) {
            Class<?> aClass = reportObjects.stream().findFirst().get().getClass();
            Workbook workbook = new XSSFWorkbook();
            Sheet sheet = workbook.createSheet(sheetName);
            int proceedRowNumber = startRowNumber;

            Row headerRow = sheet.createRow(startRowNumber);
            createHeadersFromDeclaredExcelColumns(headerRow, aClass);
            startRowNumber++;
            proceedRowNumber++;

            for (T report : reportObjects) {
                Row bodyRow = sheet.createRow(proceedRowNumber);
                createCellsFromDeclaredExcelColumns(bodyRow, report);
                proceedRowNumber++;
            }

            createTotalFormula(aClass, sheet.createRow(proceedRowNumber), startRowNumber);

            log.debug("Total rows number is: " + proceedRowNumber);
            autosizeAllByRow(headerRow);
            return workbook;
        } else {
            throw new ExcelMapperException("Couldn't get object from given list: " + reportObjects);
        }
    }

    private <T> void createHeadersFromDeclaredExcelColumns(Row row, Class<T> clazz) {
        try {
            PropertyDescriptor[] propertyDescriptors = Introspector.getBeanInfo(clazz).getPropertyDescriptors();

            for (var propertyDescriptor : propertyDescriptors)
                createHeaderFromDeclaredExcelColumns(row, clazz, propertyDescriptor);

            for (Method method : clazz.getDeclaredMethods())
                createHeaderFromDeclaredExcelFormula(row, method);

        } catch (IntrospectionException ex) {
            log.debug(ex.getLocalizedMessage());
        }
    }

    private <T> void createHeaderFromDeclaredExcelColumns(Row row, Class<T> clazz, PropertyDescriptor propertyDescriptor) {
        try {
            Field field = clazz.getDeclaredField(propertyDescriptor.getName());

            ColumnExcel columnExcel = field.getDeclaredAnnotation(ColumnExcel.class);
            if (nonNull(columnExcel)) {
                createHeader(row, columnExcel.position(), columnExcel.applyNames()[0], columnExcel.headerStyle());
            }
        } catch (NoSuchFieldException e) {
            log.debug(e.getLocalizedMessage());
        }
    }

    private void createHeaderFromDeclaredExcelFormula(Row row, Method method) {
        var columnExcelFormula = method.getDeclaredAnnotation(ColumnExcelFormula.class);

        if (nonNull(columnExcelFormula))
            createHeader(row, columnExcelFormula.position(), columnExcelFormula.name(), columnExcelFormula.headerStyle());
    }

    private void createHeader(Row row, int position, String name, ColumnExcelStyle columnExcelStyle) {
        Cell cell = row.createCell(position);
        cell.setCellValue(name);

        setCellFormatting(cell, columnExcelStyle);
        row.getSheet().autoSizeColumn(cell.getColumnIndex());
    }

    private <T> void createCellsFromDeclaredExcelColumns(Row row, T tObject) {
        try {
            PropertyDescriptor[] propertyDescriptors = Introspector.getBeanInfo(tObject.getClass()).getPropertyDescriptors();
            for (PropertyDescriptor propertyDescriptor : propertyDescriptors) {
                createCellFromDeclaredExcelColumns(row, tObject, propertyDescriptor);
            }
        } catch (IntrospectionException ex) {
            log.debug(ex.getLocalizedMessage());
        }

        for (Method method : tObject.getClass().getDeclaredMethods()) {
            try {
                createCellFromDeclaredExcelFormula(row, tObject, method);
            } catch (IllegalAccessException | InvocationTargetException ex) {
                log.debug(ex.getLocalizedMessage());
            }

        }
    }

    private <T> void createCellFromDeclaredExcelColumns(Row row, T tObject, PropertyDescriptor propertyDescriptor) {
        try {
            Field field = tObject.getClass().getDeclaredField(propertyDescriptor.getName());
            Method readMethod = propertyDescriptor.getReadMethod();

            ColumnExcel columnExcel = field.getDeclaredAnnotation(ColumnExcel.class);
            if (nonNull(columnExcel)) {
                Class<?> returnType = readMethod.getReturnType();
                Cell cell = row.createCell(columnExcel.position());

                Object invokeResult = readMethod.invoke(tObject);
                if (nonNull(invokeResult)) {
                    defineAndAssignCellValue(returnType, cell, invokeResult, readMethod);
                }
                setCellFormatting(cell, columnExcel.cellStyle());
            }
        } catch (NoSuchFieldException | InvocationTargetException | IllegalAccessException e) {
            log.debug(e.getLocalizedMessage());
        }

    }

    private void defineAndAssignCellValue(Class<?> returnType, Cell cell, Object invokeResult, Method readMethod) {
        if (returnType.isAssignableFrom(String.class)) {
            cell.setCellValue((String) invokeResult);

        } else if (returnType.isAssignableFrom(Double.class)) {
            cell.setCellValue((Double) invokeResult);

        } else if (returnType.isAssignableFrom(BigDecimal.class)) {
            cell.setCellValue(((BigDecimal) invokeResult).doubleValue());

        } else if (returnType.isAssignableFrom(Date.class)) {
            cell.setCellValue(java.util.Date.from(((Date) invokeResult).toInstant()));

        } else if (returnType.isAssignableFrom(LocalDate.class)) {
            cell.setCellValue(java.util.Date.from(((LocalDate) invokeResult).atStartOfDay(ZoneId.systemDefault()).toInstant()));

        } else if (returnType.isAssignableFrom(Integer.class)) {
            cell.setCellValue((Integer) invokeResult);

        } else if (returnType.isAssignableFrom(Long.class)) {
            cell.setCellValue(((Long) invokeResult).intValue());
        } else {
            log.debug(" Return type for the method: " + readMethod.getName() + " with @ColumnExcel annotation is not supported " +
                    "for now return type is: " + returnType.getName() + " method is ignored for the reason");
        }
    }

    @SneakyThrows
    private <T> void createTotalFormula(Class<T> tClazz, Row row, int firstRowNum) {
        Method[] methods = tClazz.getDeclaredMethods();
        for (Method method : methods) {
            ColumnExcelTotalFormula columnExcelTotalFormula = method.getAnnotation(ColumnExcelTotalFormula.class);
            if (columnExcelTotalFormula != null
                    && method.getReturnType().isAssignableFrom(String.class)
                    && method.getParameters().length == 2
                    && (method.getModifiers() & Modifier.STATIC) != 0
                    && (method.getModifiers() & Modifier.PRIVATE) == 0
            ) {
                String cellFormula = (String) method.invoke(tClazz, firstRowNum, row.getRowNum());
                Cell cell = row.createCell(columnExcelTotalFormula.position());
                cell.setCellFormula(cellFormula);

                if (columnExcelTotalFormula.useValue()) {
                    cell = applyFormulasValue(cell);
                }
                setCellFormatting(cell, columnExcelTotalFormula.cellStyle());
            }
        }
    }

    private Cell applyFormulasValue(Cell cell) {
        Row row = cell.getRow();
        int columnIndex = cell.getColumnIndex();
        FormulaEvaluator evaluator = row.getSheet().getWorkbook().getCreationHelper().createFormulaEvaluator();
        CellValue cellValue = evaluator.evaluate(cell);

        switch (cellValue.getCellType()) {
            case BOOLEAN -> {
                row.removeCell(cell);
                cell = row.createCell(columnIndex);
                cell.setCellValue(cellValue.getBooleanValue());
            }
            case NUMERIC -> {
                row.removeCell(cell);
                cell = row.createCell(columnIndex);
                cell.setCellValue(cellValue.getNumberValue());
            }
            case STRING -> {
                row.removeCell(cell);
                cell = row.createCell(columnIndex);
                cell.setCellValue(cellValue.getStringValue());
            }
            default -> {
            }
        }
        return cell;
    }

    private <T> void createCellFromDeclaredExcelFormula(Row row, T tObject, Method readMethod) throws IllegalAccessException, InvocationTargetException {
        ColumnExcelFormula columnExcelFormula = readMethod.getDeclaredAnnotation(ColumnExcelFormula.class);
        if (columnExcelFormula != null) {
            Class<?> returnType = readMethod.getReturnType();
            Cell cell = row.createCell(columnExcelFormula.position());

            if (returnType.isAssignableFrom(String.class)) {
                cell.setCellFormula((String) readMethod.invoke(tObject, row.getRowNum()));
            } else {
                log.debug(" Return type for the method: " + readMethod.getName() + " with @ColumnExcelFormula annotation has to be String " +
                        "and now it's: " + returnType.getName() + " method is ignored for the reason");
            }
            setCellFormatting(cell, columnExcelFormula.cellStyle());
        }
    }

    private void setCellFormatting(Cell cell, ColumnExcelStyle columnExcelStyle) {
        if (columnExcelStyle != null) {
            CellStyle cellStyle = cell.getRow().getSheet().getWorkbook().createCellStyle();

            Font font = cell.getRow().getSheet().getWorkbook().createFont();
            font.setBold(columnExcelStyle.isFontBold());

            if (!columnExcelStyle.fontName().equals(ExcelColumnFont.DEFAULT)) {
                font.setFontName(columnExcelStyle.fontName().getFontName());

            }
            if (!columnExcelStyle.fontColor().equals(ExcelColumnCellTextColor.AUTOMATIC)) {
                font.setColor(columnExcelStyle.fontColor().getColorIndex());
            }
            if (columnExcelStyle.fontSize() != -1) {
                font.setFontHeightInPoints(columnExcelStyle.fontSize());
            }

            if (!columnExcelStyle.cellTypePattern().equals(ExcelColumnDataFormat.NONE)) {
                DataFormat dataFormat = cell.getRow().getSheet().getWorkbook().createDataFormat();
                cellStyle.setDataFormat(dataFormat.getFormat(columnExcelStyle.cellTypePattern().getFormatPattern()));

            }
            if (!columnExcelStyle.cellColor().equals(IndexedColors.AUTOMATIC)) {
                cellStyle.setFillForegroundColor(columnExcelStyle.cellColor().getIndex());
                cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            }
            if (columnExcelStyle.isCentreAlignment()) {
                cellStyle.setAlignment(HorizontalAlignment.CENTER);
                cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);

            }
            if (columnExcelStyle.isFramed()) {
                cellStyle.setBorderTop(BorderStyle.THIN);
                cellStyle.setBorderBottom(BorderStyle.THIN);
                cellStyle.setBorderLeft(BorderStyle.THIN);
                cellStyle.setBorderRight(BorderStyle.THIN);
            }

            cellStyle.setFont(font);
            cell.setCellStyle(cellStyle);
        }
    }

    private void autosizeAllByRow(Row row) {
        for (int i = row.getFirstCellNum(); i <= row.getLastCellNum(); i++) {
            row.getSheet().autoSizeColumn(i);
        }
    }
}

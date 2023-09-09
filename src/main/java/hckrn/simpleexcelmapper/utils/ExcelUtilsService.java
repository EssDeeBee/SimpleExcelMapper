package hckrn.simpleexcelmapper.utils;


import hckrn.simpleexcelmapper.annotation.ColumnExcel;
import hckrn.simpleexcelmapper.annotation.ColumnExcelFormula;
import hckrn.simpleexcelmapper.annotation.ColumnExcelStyle;
import hckrn.simpleexcelmapper.annotation.ColumnExcelTotalFormula;
import hckrn.simpleexcelmapper.exception.ExcelHeaderNotFoundException;
import hckrn.simpleexcelmapper.format.ExcelColumnCellTextColor;
import hckrn.simpleexcelmapper.format.ExcelColumnDataFormat;
import hckrn.simpleexcelmapper.format.ExcelColumnFont;
import lombok.SneakyThrows;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.beans.IntrospectionException;
import java.beans.Introspector;
import java.beans.PropertyDescriptor;
import java.lang.reflect.*;
import java.math.BigDecimal;
import java.sql.Date;
import java.text.NumberFormat;
import java.time.LocalDate;
import java.time.ZoneId;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;

@Slf4j
public class ExcelUtilsService {

    @SneakyThrows
    public <T> List<T> readTableFromSheet(Sheet sheet, Class<T> clazz) {
        List<T> tArrayList = new ArrayList<>();

        ExcelHeaders excelHeaders = findHeaders(clazz, sheet);
        int dataRowNumber = excelHeaders.getHeadersRowNumber() + 1;
        int endRowNum = sheet.getLastRowNum();

        for (; dataRowNumber <= endRowNum; dataRowNumber++) {
            tArrayList.add(createObjectFromDeclaredExcelColumns(sheet.getRow(dataRowNumber), clazz, excelHeaders));
        }
        return tArrayList;
    }

    public <T> T createObjectFromDeclaredExcelColumns(Row row, Class<T> reportClass, ExcelHeaders excelHeaders)
            throws
            IllegalAccessException,
            InvocationTargetException,
            IntrospectionException {

        T tObject = createNewEntityInstance(reportClass);
        HashMap<String, Integer> excelIndexes = excelHeaders.getHeadersIndexes();

        for (PropertyDescriptor propertyDescriptor : Introspector.getBeanInfo(tObject.getClass()).getPropertyDescriptors()) {
            Integer cellIndex = excelIndexes.get(propertyDescriptor.getName());

            if (cellIndex != null) {
                Method writeMethod = propertyDescriptor.getWriteMethod();
                Parameter[] parameters = writeMethod.getParameters();

                if (parameters.length == 1) {
                    Cell cell = row.getCell(cellIndex);
                    if (cell != null) {
                        if (parameters[0].getType().isAssignableFrom(String.class)) {
                            writeMethod.invoke(tObject, getStringCellValueOrNull(cell));

                        } else if (parameters[0].getType().isAssignableFrom(Double.class)) {
                            writeMethod.invoke(tObject, getDoubleCellValueOrNullFromMergedCells(cell));

                        } else if (parameters[0].getType().isAssignableFrom(Date.class)) {
                            writeMethod.invoke(tObject, getDateCellOrNull(cell));

                        } else if (!parameters[0].getType().isPrimitive() && parameters[0].getType().isAssignableFrom(Integer.class)) {
                            writeMethod.invoke(tObject, getDoubleCellValueOrNull(cell) == null ? null : getDoubleCellValueOrNull(cell).intValue());

                        } else if (parameters[0].getType().isPrimitive() && parameters[0].getType().isAssignableFrom(Integer.class)) {
                            writeMethod.invoke(tObject, getDoubleCellValueOrNull(cell) == null ? 0 : getDoubleCellValueOrNull(cell).intValue());
                        }
                    }
                } else {
                    log.debug(writeMethod.getName() + " method ignored while creating entity param, reason: method has not less or more than one parameter");
                }
            }
        }
        return tObject;
    }

    public final Double getDoubleCellValueOrNull(Cell cell) {
        Double cellValue = null;
        if (cell != null) {
            if (cell.getCellType().equals(CellType.NUMERIC) || cell.getCellType().equals(CellType.FORMULA)) {
                try {
                    cellValue = cell.getNumericCellValue();
                } catch (NumberFormatException | IllegalStateException ex) {
                    log.debug(ex.getMessage());
                }
            }
        }
        return cellValue;
    }

    public final Double getDoubleCellValueOrNullFromMergedCells(Cell cell) {
        if (cell != null) {
            List<CellRangeAddress> cellAddresses = cell.getRow().getSheet().getMergedRegions();

            for (CellRangeAddress cellAddress : cellAddresses) {
                if (cellAddress.isInRange(cell)) {
                    Cell firstMergedCell = cell.getRow().getSheet().getRow(cellAddress.getFirstRow()).getCell(cellAddress.getFirstColumn());

                    return getDoubleCellValueOrNull(firstMergedCell);
                }
            }
        }
        return getDoubleCellValueOrNull(cell);
    }

    public final String getStringCellValueOrNull(Cell cell) {
        String cellValue = null;
        if (cell != null) {

            if (cell.getCellType().equals(CellType.STRING) || cell.getCellType().equals(CellType.FORMULA)) {
                try {
                    cellValue = cell.getStringCellValue();
                } catch (NumberFormatException | IllegalStateException ex) {
                    log.debug(ex.getMessage());
                }
            } else if (cell.getCellType().equals(CellType.NUMERIC)) {
                try {
                    //TODO check other options to format
                    NumberFormat fmt = NumberFormat.getInstance();
                    fmt.setGroupingUsed(false);
                    fmt.setMaximumIntegerDigits(999);
                    fmt.setMaximumFractionDigits(999);

                    double numericCellValue = cell.getNumericCellValue();
                    cellValue = fmt.format(numericCellValue);

                } catch (NumberFormatException | IllegalStateException ex) {
                    log.debug(ex.getMessage());
                }
            }
        }
        return cellValue;


    }

    public final Date getDateCellOrNull(Cell cell) {
        if (cell != null) {
            if (cell.getCellType().equals(CellType.NUMERIC) || cell.getCellType().equals(CellType.STRING) || cell.getCellType().equals(CellType.FORMULA)) {
                try {
                    java.util.Date cellValue = cell.getDateCellValue();
                    return Date.valueOf(cellValue.toInstant().atZone(ZoneId.systemDefault()).toLocalDate());

                } catch (NumberFormatException | IllegalStateException ex) {
                    log.debug(ex.getMessage());
                }
            }
        }
        return null;
    }

    protected <T> T createNewEntityInstance(Class<T> t) {
        try {
            Constructor[] constructors = t.getConstructors();
            for (Constructor constructor : constructors) {
                if (constructor.getParameterCount() == 0) {
                    return (T) constructor.newInstance();
                }
            }

        } catch (IllegalAccessException | InvocationTargetException | InstantiationException ex) {
            throw new RuntimeException(ex);
        }
        throw new RuntimeException();
    }

    public final <T> void createHeadersFromDeclaredExcelColumns(Row row, Class<T> clazz) {
        try {
            PropertyDescriptor[] propertyDescriptors = Introspector.getBeanInfo(clazz).getPropertyDescriptors();

            for (PropertyDescriptor propertyDescriptor : propertyDescriptors) {
                try {
                    createHeaderFromDeclaredExcelColumns(row, clazz, propertyDescriptor);
                } catch (NoSuchFieldException e) {
                    log.debug(e.getLocalizedMessage());
                }

            }

            for (Method method : clazz.getDeclaredMethods())
                createHeaderFromDeclaredExcelFormula(row, method);

        } catch (IntrospectionException ex) {
            log.debug(ex.getLocalizedMessage());
        }
    }

    protected <T> void createHeaderFromDeclaredExcelColumns(Row row, Class<T> clazz, PropertyDescriptor propertyDescriptor) throws NoSuchFieldException {
        Field field = clazz.getDeclaredField(propertyDescriptor.getName());

        ColumnExcel columnExcel = field.getDeclaredAnnotation(ColumnExcel.class);
        if (columnExcel != null) {
            createHeader(row, columnExcel.position(), columnExcel.applyNames()[0], columnExcel.headerStyle());
        }


    }

    protected void createHeaderFromDeclaredExcelFormula(Row row, Method method) {
        ColumnExcelFormula columnExcelFormula = method.getDeclaredAnnotation(ColumnExcelFormula.class);

        if (columnExcelFormula != null) {
            createHeader(row, columnExcelFormula.position(), columnExcelFormula.name(), columnExcelFormula.headerStyle());
        }
    }

    protected void createHeader(Row row, int position, String name, ColumnExcelStyle columnExcelStyle) {
        Cell cell = row.createCell(position);
        cell.setCellValue(name);

        setCellFormatting(cell, columnExcelStyle);
        row.getSheet().autoSizeColumn(cell.getColumnIndex());
    }

    public final <T> void createCellsFromDeclaredExcelColumns(Row row, T tObject) {
        try {
            PropertyDescriptor[] propertyDescriptors = Introspector.getBeanInfo(tObject.getClass()).getPropertyDescriptors();

            for (PropertyDescriptor propertyDescriptor : propertyDescriptors) {
                try {
                    createCellFromDeclaredExcelColumns(row, tObject, propertyDescriptor);
                } catch (NoSuchFieldException | InvocationTargetException | IllegalAccessException e) {
                    log.debug(e.getLocalizedMessage());
                }

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


    protected <T> void createCellFromDeclaredExcelColumns(Row row, T tObject, PropertyDescriptor propertyDescriptor) throws NoSuchFieldException, InvocationTargetException, IllegalAccessException {
        Field field = tObject.getClass().getDeclaredField(propertyDescriptor.getName());
        Method readMethod = propertyDescriptor.getReadMethod();

        ColumnExcel columnExcel = field.getDeclaredAnnotation(ColumnExcel.class);
        if (columnExcel != null) {
            Class<?> returnType = readMethod.getReturnType();
            Cell cell = row.createCell(columnExcel.position());

            if (returnType != null) {
                Object invokeResult = readMethod.invoke(tObject);
                if (invokeResult != null) {

                    if (returnType.isAssignableFrom(String.class)) {
                        cell.setCellValue((String) invokeResult);

                    } else if (returnType.isAssignableFrom(Double.class)) {
                        cell.setCellValue((Double) invokeResult);

                    } else if (returnType.isAssignableFrom(BigDecimal.class)) {
                        cell.setCellValue(((BigDecimal) invokeResult).doubleValue());

                    } else if (returnType.isAssignableFrom(Date.class)) {
                        cell.setCellValue(java.util.Date.from(((Date) invokeResult).toInstant()));

                    } else if (returnType.isAssignableFrom(Integer.class)) {
                        cell.setCellValue((Integer) invokeResult);

                    } else if (returnType.isAssignableFrom(Long.class)) {
                        cell.setCellValue(((Long) invokeResult).intValue());

                    } else {
                        log.debug(" Return type for the method: " + readMethod.getName() + " with @ColumnExcel annotation is not supported " +
                                "for now return type is: " + returnType.getName() + " method is ignored for the reason");
                    }
                }
            }
            setCellFormatting(cell, columnExcel.cellStyle());
        }
    }

    @SneakyThrows
    public <T> void createTotalFormula(Class<T> tClazz, Row row, int firstRowNum) {
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
            case BOOLEAN:
                row.removeCell(cell);
                cell = row.createCell(columnIndex);
                cell.setCellValue(cellValue.getBooleanValue());
                break;
            case NUMERIC:
                row.removeCell(cell);
                cell = row.createCell(columnIndex);
                cell.setCellValue(cellValue.getNumberValue());
                break;
            case STRING:
                row.removeCell(cell);
                cell = row.createCell(columnIndex);
                cell.setCellValue(cellValue.getStringValue());
                break;
            case BLANK:
                break;
            case ERROR:
                break;
            case FORMULA:
                break;
        }
        return cell;
    }

    protected <T> void createCellFromDeclaredExcelFormula(Row row, T tObject, Method readMethod) throws IllegalAccessException, InvocationTargetException {
        ColumnExcelFormula columnExcelFormula = readMethod.getDeclaredAnnotation(ColumnExcelFormula.class);
        if (columnExcelFormula != null) {
            Class<?> returnType = readMethod.getReturnType();
            Cell cell = row.createCell(columnExcelFormula.position());

            if (returnType != null) {
                if (returnType.isAssignableFrom(String.class)) {
                    cell.setCellFormula((String) readMethod.invoke(tObject, row.getRowNum()));
                } else {
                    log.debug(" Return type for the method: " + readMethod.getName() + " with @ColumnExcelFormula annotation has to be String " +
                            "and now it's: " + returnType.getName() + " method is ignored for the reason");
                }
            }
            setCellFormatting(cell, columnExcelFormula.cellStyle());
        }
    }

    protected void setCellFormatting(Cell cell, ColumnExcelStyle columnExcelStyle) {
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

    public <T> ExcelHeaders findHeaders(Class<T> clazz, Sheet sheet) {
        return findHeaders(clazz, sheet, sheet.getLastRowNum());
    }

    public <T> ExcelHeaders findHeaders(Class<T> clazz, Sheet sheet, int rowsForCheck) {

        if (rowsForCheck <= 0) {
            throw new RuntimeException("Number rows for checking cannot be less or equals zero");
        }

        try {
            HashMap<String, List<String>> fieldsWithoutIndex = new HashMap<>();
            HashMap<String, Integer> fieldsToIndex = new HashMap<>();
            PropertyDescriptor[] propertyDescriptors = Introspector.getBeanInfo(clazz).getPropertyDescriptors();
            List<String> uniqueApplyNames = new ArrayList<>(propertyDescriptors.length * 3);

            forExcelRows:
            for (int rowNum = 0; rowNum < rowsForCheck; rowNum++) {
                Row headerRow = sheet.getRow(rowNum);
                if (headerRow != null) {
                    fieldsWithoutIndex.clear();
                    fieldsToIndex.clear();

                    forApplyNames:
                    for (PropertyDescriptor propertyDescriptor : propertyDescriptors) {
                        String propertyName = propertyDescriptor.getName();
                        uniqueApplyNames.clear();


                        try {
                            Field declaredField = clazz.getDeclaredField(propertyName);

                            //Getting annotated fields
                            if (declaredField != null) {
                                ColumnExcel excelColumn = declaredField.getDeclaredAnnotation(ColumnExcel.class);

                                if (excelColumn != null) {
                                    String[] applyNames = excelColumn.applyNames();

                                    //Getting indexes for annotated fields
                                    for (int j = 0; j < applyNames.length; j++) {

                                        String applyNameForComparing = applyNames[j]
                                                .replace(" ", "")
                                                .toUpperCase()
                                                .strip();

                                        if (uniqueApplyNames.contains(applyNameForComparing)) {
                                            throw new RuntimeException("Duplicate header was found: " + applyNames[j]);
                                        } else {
                                            uniqueApplyNames.add(applyNameForComparing);
                                        }

                                        for (Cell cell : sheet.getRow(rowNum)) {
                                            String cellValue = getStringCellValueOrNull(cell);
                                            if (cellValue != null
                                                    && applyNameForComparing.equalsIgnoreCase(cellValue
                                                    .replace(" ", "")
                                                    .replace("\n", "")
                                                    .toUpperCase()
                                                    .strip())) {
                                                if (fieldsToIndex.get(propertyName) != null)
                                                    throw new RuntimeException("Duplicate header was found: " + propertyName);

                                                fieldsToIndex.put(propertyName, cell.getColumnIndex());
                                                continue forApplyNames;
                                            }
                                        }

                                        if (j == (applyNames.length - 1)) {
                                            log.debug("Headers not found. Row: " + rowNum + " Header: " + applyNames[0] + "\nSwitching to the next row");
                                            fieldsWithoutIndex.put(propertyName, List.of(applyNames));
                                            continue forExcelRows;
                                        }
                                    }
                                }
                            }
                        } catch (NoSuchFieldException ex) {
                            ex.getMessage();
                        }

                    }
                }
                if (fieldsWithoutIndex.isEmpty() && !fieldsToIndex.isEmpty()) {
                    return new ExcelHeaders(fieldsToIndex, rowNum);

                }
            }

            log.error("Headers are not found! Headers were looking for class: " + clazz.getName()
                    + " on sheet: " + sheet.getSheetName()
                    + " for first: " + rowsForCheck + " rows");
            throw new ExcelHeaderNotFoundException(null, "", "for first: " + rowsForCheck + " rows");

        } catch (IntrospectionException ex) {
            throw new RuntimeException(ex);
        }
    }

    public <T> Workbook createWorkbookFromObject(List<T> reportObjects) {
        return createWorkbookFromObject(reportObjects, 0, "Report_" + LocalDate.now().toString());
    }

    public <T> Workbook createReportWorkbook(List<T> reportObjects, int startRowNumber) {
        return createWorkbookFromObject(reportObjects, startRowNumber, "Report_" + LocalDate.now().toString());
    }

    public <T> Workbook createWorkbookFromObject(List<T> reportObjects, int startRowNumber, String sheetName) {

        if (reportObjects.stream().findFirst().isPresent()) {
            T obj = reportObjects.stream().findFirst().get();
            Workbook workbook = new XSSFWorkbook();
            Sheet sheet = workbook.createSheet(sheetName);
            int proceedRowNumber = startRowNumber;

            Row headerRow = sheet.createRow(startRowNumber);
            createHeadersFromDeclaredExcelColumns(headerRow, obj.getClass());
            startRowNumber++;
            proceedRowNumber++;

            for (T report : reportObjects) {
                Row bodyRow = sheet.createRow(proceedRowNumber);
                createCellsFromDeclaredExcelColumns(bodyRow, report);
                proceedRowNumber++;
            }

            createTotalFormula(obj.getClass(), sheet.createRow(proceedRowNumber), startRowNumber);

            log.info("Total rows number is: " + proceedRowNumber);
            autosizeAllByRow(headerRow);
            return workbook;
        } else {
            throw new RuntimeException("Couldn't get object from given list: " + reportObjects);
        }
    }

    private void autosizeAllByRow(Row row) {
        for (int i = row.getFirstCellNum(); i <= row.getLastCellNum(); i++) {
            row.getSheet().autoSizeColumn(i);
        }
    }
}

package hckrn.simpleexcelmapper.utils;


import hckrn.simpleexcelmapper.annotation.ColumnExcel;
import lombok.SneakyThrows;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.beans.Introspector;
import java.beans.PropertyDescriptor;
import java.lang.reflect.Constructor;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.lang.reflect.Parameter;
import java.sql.Date;
import java.time.ZoneId;
import java.util.*;


@Slf4j
public class ExcelMapper {


    public final <T> Map<String, List<T>> mapWorkbookToObjs(Workbook workbook, Class<T> tObject) {
        var result = new HashMap<String, List<T>>();
        Iterator<Sheet> sheetIterator = workbook.sheetIterator();
        while (sheetIterator.hasNext()) {
            Sheet sheet = sheetIterator.next();
            result.put(sheet.getSheetName(), mapSheetToObj(sheet, tObject, true));
        }
        return result;
    }

    public final <T> List<T> mapSheetToObj(Sheet sheet, Class<T> tObject, boolean ignoreFirstRow) {
        var result = new LinkedList<T>();
        Iterator<Row> rowIterator = sheet.rowIterator();
        if (rowIterator.hasNext() && ignoreFirstRow) {
            rowIterator.next();
        }
        while (rowIterator.hasNext()) {
            result.add(createObjectFromDeclaredExcelColumns(rowIterator.next(), tObject));
        }
        return result;
    }


    @SneakyThrows
    public final <T> T createObjectFromDeclaredExcelColumns(Row row, Class<T> tClass) {
        Constructor<?>[] constructors = tClass.getConstructors();
        Constructor<?> noArgsConstructor = Arrays.stream(constructors).filter(constructor -> constructor.getParameterCount() == 0).findFirst().orElseThrow(RuntimeException::new);
        T toObject = (T) noArgsConstructor.newInstance();

        PropertyDescriptor[] propertyDescriptors = Introspector.getBeanInfo(tClass).getPropertyDescriptors();
        for (PropertyDescriptor propertyDescriptor : propertyDescriptors) {

            try {
                Field field = tClass.getDeclaredField(propertyDescriptor.getName());
                ColumnExcel columnExcel = field.getDeclaredAnnotation(ColumnExcel.class);
                Method writeMethod = propertyDescriptor.getWriteMethod();
                if (columnExcel != null) {
                    Parameter[] parameters = writeMethod.getParameters();

                    if (parameters.length == 1) {
                        Cell cell = row.getCell(columnExcel.position());
                        if (cell != null) {
                            if (parameters[0].getType().isAssignableFrom(String.class)) {
                                writeMethod.invoke(toObject, getStringCellValueOrEmpty(cell));

                            } else if (parameters[0].getType().isAssignableFrom(Double.class)) {
                                writeMethod.invoke(toObject, getDoubleCellValueOrZeroFromMergedCells(cell));

                            } else if (parameters[0].getType().isAssignableFrom(Date.class)) {
                                writeMethod.invoke(toObject, getDateCellOrNull(cell));

                            } else if (parameters[0].getType().isAssignableFrom(Integer.class)) {
                                writeMethod.invoke(toObject, getDoubleCellValueOrZero(cell).intValue());
                            }
                        }
                    } else {
                        log.debug(writeMethod.getName() + " method ignored while creating entity param, reason: method has not less or more than one parameter");
                    }
                }
            } catch (NoSuchFieldException ex) {
                log.debug("");
            }

        }
        return toObject;
    }

    private double getDoubleCellValueOrZeroFromMergedCells(Cell cell) {
        if (cell != null) {
            List<CellRangeAddress> cellAddresses = cell.getRow().getSheet().getMergedRegions();

            for (CellRangeAddress cellAddress : cellAddresses) {
                if (cellAddress.isInRange(cell)) {
                    Cell firstMergedCell = cell.getRow().getSheet().getRow(cellAddress.getFirstRow()).getCell(cellAddress.getFirstColumn());
                    return getDoubleCellValueOrZero(firstMergedCell);
                }
            }
        }
        return getDoubleCellValueOrZero(cell);
    }

    private String getStringCellValueOrEmpty(Cell cell) {
        String cellValue = "";
        if (cell != null && (cell.getCellType().equals(CellType.STRING) || cell.getCellType().equals(CellType.FORMULA))) {
            try {
                cellValue = cell.getStringCellValue();
            } catch (NumberFormatException | IllegalStateException ex) {
                log.debug(ex.getMessage());
            }

        }
        return cellValue;
    }

    private Double getDoubleCellValueOrZero(Cell cell) {
        double cellValue = 0d;
        if (cell != null && (cell.getCellType().equals(CellType.NUMERIC) || cell.getCellType().equals(CellType.FORMULA))) {
            try {
                cellValue = cell.getNumericCellValue();
            } catch (NumberFormatException | IllegalStateException ex) {
                log.debug(ex.getMessage());
            }

        }
        return cellValue;
    }

    private Date getDateCellOrNull(Cell cell) {
        if (cell != null && (cell.getCellType().equals(CellType.NUMERIC) || cell.getCellType().equals(CellType.STRING) || cell.getCellType().equals(CellType.FORMULA))) {
            try {
                java.util.Date cellValue = cell.getDateCellValue();
                return Date.valueOf(cellValue.toInstant().atZone(ZoneId.systemDefault()).toLocalDate());

            } catch (NumberFormatException | IllegalStateException ex) {
                log.debug(ex.getMessage());
            }

        }
        return null;
    }
}

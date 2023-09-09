package hckrn.simpleexcelmapper.utils;


import hckrn.simpleexcelmapper.annotation.ColumnExcel;
import lombok.SneakyThrows;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.beans.Introspector;
import java.beans.PropertyDescriptor;
import java.lang.reflect.*;
import java.sql.Date;
import java.time.LocalDate;
import java.time.ZoneId;
import java.util.*;

import static java.util.Objects.nonNull;
import static org.apache.poi.ss.usermodel.CellType.*;


@Slf4j
class ExcelToObjMapper {

    <T> Map<String, List<T>> mapWorkbookToObjs(Workbook workbook, Class<T> tObject) {
        var result = new HashMap<String, List<T>>();
        Iterator<Sheet> sheetIterator = workbook.sheetIterator();
        while (sheetIterator.hasNext()) {
            Sheet sheet = sheetIterator.next();
            result.put(sheet.getSheetName(), mapSheetToObjs(sheet, tObject, true));
        }
        return result;
    }

    <T> List<T> mapSheetToObjs(Sheet sheet, Class<T> tObject, boolean ignoreFirstRow) {
        var result = new LinkedList<T>();
        Iterator<Row> rowIterator = sheet.rowIterator();
        if (rowIterator.hasNext() && ignoreFirstRow) {
            rowIterator.next();
        }
        while (rowIterator.hasNext()) {
            result.add(mapRowToObj(rowIterator.next(), tObject));
        }
        return result;
    }


    @SneakyThrows
    <T> T mapRowToObj(Row row, Class<T> tClass) {
        Constructor<?>[] constructors = tClass.getConstructors();
        Constructor<?> noArgsConstructor = Arrays.stream(constructors)
                .filter(constructor -> constructor.getParameterCount() == 0).findFirst().orElseThrow(RuntimeException::new);
        T toObject = (T) noArgsConstructor.newInstance();

        PropertyDescriptor[] propertyDescriptors = Introspector.getBeanInfo(tClass).getPropertyDescriptors();
        for (var propertyDescriptor : propertyDescriptors) {

            try {
                Field field = tClass.getDeclaredField(propertyDescriptor.getName());
                ColumnExcel columnExcel = field.getDeclaredAnnotation(ColumnExcel.class);
                Method writeMethod = propertyDescriptor.getWriteMethod();
                if (columnExcel != null) {
                    Parameter[] parameters = writeMethod.getParameters();
                    if (parameters.length == 1) {
                        Cell cell = row.getCell(columnExcel.position());
                        if (nonNull(cell)) {
                            extractAndWriteValue(parameters, writeMethod, toObject, cell);
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

    private <T> void extractAndWriteValue(Parameter[] parameters, Method writeMethod, T toObject, Cell cell) throws IllegalAccessException, InvocationTargetException {
        Class<?> type = parameters[0].getType();

        if (type.isAssignableFrom(String.class)) {
            writeMethod.invoke(toObject, getStringCellValueOrEmpty(cell));

        } else if (type.isAssignableFrom(Double.class)) {
            writeMethod.invoke(toObject, getDoubleCellValueOrZeroFromMergedCells(cell));

        } else if (type.isAssignableFrom(Date.class)) {
            writeMethod.invoke(toObject, getDateCellOrNull(cell));

        } else if (type.isAssignableFrom(LocalDate.class)) {
            writeMethod.invoke(toObject, getLocalDateCellOrNull(cell));

        } else if (type.isAssignableFrom(Integer.class)) {
            writeMethod.invoke(toObject, getDoubleCellValueOrZero(cell).intValue());
        }
    }

    private double getDoubleCellValueOrZeroFromMergedCells(Cell cell) {
        if (cell != null) {
            List<CellRangeAddress> cellAddresses = cell.getRow().getSheet().getMergedRegions();

            for (var cellAddress : cellAddresses) {
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
        if (nonNull(cell)) {
            CellType cellType = cell.getCellType();
            if (STRING.equals(cellType) || FORMULA.equals(cellType)) {
                try {
                    cellValue = cell.getStringCellValue();
                } catch (NumberFormatException | IllegalStateException ex) {
                    log.debug(ex.getMessage());
                }
            }
        }
        return cellValue;
    }

    private Double getDoubleCellValueOrZero(Cell cell) {
        double cellValue = 0d;
        if (nonNull(cell)) {
            CellType cellType = cell.getCellType();
            if (NUMERIC.equals(cellType) || FORMULA.equals(cellType)) {
                try {
                    cellValue = cell.getNumericCellValue();
                } catch (NumberFormatException | IllegalStateException ex) {
                    log.debug(ex.getMessage());
                }
            }
        }
        return cellValue;
    }

    private Date getDateCellOrNull(Cell cell) {
        Date date = null;
        if (nonNull(cell)) {
            CellType cellType = cell.getCellType();
            if (NUMERIC.equals(cellType) || STRING.equals(cellType) || FORMULA.equals(cellType)) {
                try {
                    java.util.Date cellValue = cell.getDateCellValue();
                    date = Date.valueOf(cellValue.toInstant().atZone(ZoneId.systemDefault()).toLocalDate());

                } catch (NumberFormatException | IllegalStateException ex) {
                    log.debug(ex.getMessage());
                }
            }
        }
        return date;
    }

    private LocalDate getLocalDateCellOrNull(Cell cell) {
        LocalDate localDate = null;
        if (nonNull(cell)) {
            CellType cellType = cell.getCellType();
            if (NUMERIC.equals(cellType) || STRING.equals(cellType) || FORMULA.equals(cellType)) {
                try {
                    java.util.Date cellValue = cell.getDateCellValue();
                    localDate = cellValue.toInstant().atZone(ZoneId.systemDefault()).toLocalDate();

                } catch (NumberFormatException | IllegalStateException ex) {
                    log.debug(ex.getMessage());
                }
            }
        }
        return localDate;
    }
}

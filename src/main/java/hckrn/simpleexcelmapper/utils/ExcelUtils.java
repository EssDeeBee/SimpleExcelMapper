package hckrn.simpleexcelmapper.utils;

import hckrn.simpleexcelmapper.annotation.ColumnExcel;
import hckrn.simpleexcelmapper.annotation.ColumnExcelFormula;
import hckrn.simpleexcelmapper.annotation.ColumnExcelStyle;
import hckrn.simpleexcelmapper.format.ExcelColumnCellTextColor;
import hckrn.simpleexcelmapper.format.ExcelColumnDataFormat;
import hckrn.simpleexcelmapper.format.ExcelColumnFont;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;

import java.beans.IntrospectionException;
import java.beans.Introspector;
import java.beans.PropertyDescriptor;
import java.lang.reflect.Constructor;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.math.BigDecimal;
import java.sql.Date;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Objects;
import java.util.stream.Collectors;

@Slf4j
public class ExcelUtils {

    public final List<Integer> getListIdsFromStringCell(Row row, int columnNum) {
        CellType cellType = row.getCell(columnNum).getCellType();
        List<Integer> placementIdList = new ArrayList<>();

        if (cellType.equals(CellType.STRING)) {
            String cellStringValue = row
                    .getCell(columnNum)
                    .getStringCellValue()
                    .replace(" ", "");

            if (cellStringValue.length() > 1) {
                try {
                    placementIdList = Arrays
                            .stream(cellStringValue.split(","))
                            .map(Integer::parseInt)
                            .collect(Collectors.toList());
                } catch (NumberFormatException ex) {
                    log.debug(ex.getLocalizedMessage());
                }
            }

        } else if (cellType.equals(CellType.NUMERIC)) {
            placementIdList.add((int) row.getCell(columnNum).getNumericCellValue());
        }
        return placementIdList;
    }

    private <T> T createNewEntityInstance(Class<T> t) {
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

    private <T> void createHeaderFromDeclaredExcelColumns(Row row, Class<T> clazz, PropertyDescriptor propertyDescriptor) throws NoSuchFieldException {
        Field field = clazz.getDeclaredField(propertyDescriptor.getName());

        ColumnExcel columnExcel = field.getDeclaredAnnotation(ColumnExcel.class);
        if (columnExcel != null) {
            Cell cell = row.createCell(columnExcel.position());
            cell.setCellValue(columnExcel.name());

            setCellFormatting(cell, columnExcel.headerStyle());

            row.getSheet().autoSizeColumn(cell.getColumnIndex());
        }


    }

    private void createHeaderFromDeclaredExcelFormula(Row row, Method method) {
        ColumnExcelFormula columnExcelFormula = method.getDeclaredAnnotation(ColumnExcelFormula.class);

        if (columnExcelFormula != null) {
            Cell cell = row.createCell(columnExcelFormula.position());
            cell.setCellValue(columnExcelFormula.name());

            setCellFormatting(cell, columnExcelFormula.headerStyle());
        }
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


    private <T> void createCellFromDeclaredExcelColumns(Row row, T tObject, PropertyDescriptor propertyDescriptor) throws NoSuchFieldException, InvocationTargetException, IllegalAccessException {
        Field field = tObject.getClass().getDeclaredField(propertyDescriptor.getName());
        Method readMethod = propertyDescriptor.getReadMethod();

        ColumnExcel columnExcel = field.getDeclaredAnnotation(ColumnExcel.class);
        if (columnExcel != null) {
            Class<?> returnType = readMethod.getReturnType();
            Cell cell = row.createCell(columnExcel.position());

            if (Objects.nonNull(returnType)) {
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
                    } else {
                        log.debug(" Return type for the method: " + readMethod.getName() + " with @ColumnExcel annotation is not supported " +
                                "for now return type is: " + returnType.getName() + " method is ignored for the reason");
                    }
                }
            }
            setCellFormatting(cell, columnExcel.cellStyle());
        }
    }

    private <T> void createCellFromDeclaredExcelFormula(Row row, T tObject, Method readMethod) throws IllegalAccessException, InvocationTargetException {
        ColumnExcelFormula columnExcelFormula = readMethod.getDeclaredAnnotation(ColumnExcelFormula.class);
        if (columnExcelFormula != null) {
            Class<?> returnType = readMethod.getReturnType();
            Cell cell = row.createCell(columnExcelFormula.position());

            if (Objects.nonNull(returnType)) {
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
}

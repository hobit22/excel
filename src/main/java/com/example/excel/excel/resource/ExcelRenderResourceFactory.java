package com.example.excel.excel.resource;

import com.example.excel.excel.style.*;
import org.apache.poi.ss.usermodel.Workbook;

import java.lang.annotation.Annotation;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import static com.example.excel.excel.util.SuperClassReflectionUtils.getAllFields;
import static com.example.excel.excel.util.SuperClassReflectionUtils.getAnnotation;

public class ExcelRenderResourceFactory {



    public static <T> ExcelRenderResource prepareRenderResource(Class<?> type, Workbook wb,
                                                            DataFormatDecider dataFormatDecider, List<T> data) {
        PreCalculatedCellStyleMap styleMap = new PreCalculatedCellStyleMap(dataFormatDecider);
        Map<String, String> headerNamesMap = new LinkedHashMap<>();
        List<String> fieldNames = new ArrayList<>();

        ExcelColumnStyle classDefinedHeaderStyle = getHeaderExcelColumnStyle(type);
        ExcelColumnStyle classDefinedBodyStyle = getBodyExcelColumnStyle(type);

        for (Field field : getAllFields(type)) {
            if (field.isAnnotationPresent(ExcelColumn.class)) {
                ExcelColumn annotation = field.getAnnotation(ExcelColumn.class);
                styleMap.put(
                        String.class,
                        ExcelCellKey.of(field.getName(), ExcelRenderLocation.HEADER),
                        getCellStyle(decideAppliedStyleAnnotation(classDefinedHeaderStyle, annotation.headerStyle())), wb);
                Class<?> fieldType = field.getType();
                styleMap.put(
                        fieldType,
                        ExcelCellKey.of(field.getName(), ExcelRenderLocation.BODY),
                        getCellStyle(decideAppliedStyleAnnotation(classDefinedBodyStyle, annotation.bodyStyle())), wb);
                fieldNames.add(field.getName());
                if (field.getType().equals(List.class)) {
                    int size = 0;
                    try {
                        size = ((List<?>) field.get(data.get(0))).size();
                    } catch (IllegalAccessException e) {
                        e.printStackTrace();
                    }

                    for (int i = 0; i < size; i++) {
                        headerNamesMap.put(field.getName() + "." + i, Integer.toString(i + 1));
                    }
                } else {
                    // Otherwise, add the header name as specified in the @ExcelColumn annotation
                    headerNamesMap.put(field.getName(), annotation.headerName());
                }
//                headerNamesMap.put(field.getName(), annotation.headerName());
            }
        }

        if (styleMap.isEmpty()) {
//            throw new CustomException(EXCEL_RENDER_FAILED);
        }
        return new ExcelRenderResource(styleMap, headerNamesMap, fieldNames);
    }

    private static ExcelColumnStyle getHeaderExcelColumnStyle(Class<?> clazz) {
        Annotation annotation = getAnnotation(clazz, DefaultHeaderStyle.class);
        if (annotation == null) {
            return null;
        }
        return ((DefaultHeaderStyle) annotation).style();
    }

    private static ExcelColumnStyle getBodyExcelColumnStyle(Class<?> clazz) {
        Annotation annotation = getAnnotation(clazz, DefaultBodyStyle.class);
        if (annotation == null) {
            return null;
        }
        return ((DefaultBodyStyle) annotation).style();
    }

    private static ExcelColumnStyle decideAppliedStyleAnnotation(ExcelColumnStyle classAnnotation,
                                                                 ExcelColumnStyle fieldAnnotation) {
        if (fieldAnnotation.excelCellStyleClass().equals(NoExcelCellStyle.class) && classAnnotation != null) {
            return classAnnotation;
        }
        return fieldAnnotation;
    }

    private static ExcelCellStyle getCellStyle(ExcelColumnStyle excelColumnStyle) {
        Class<? extends ExcelCellStyle> excelCellStyleClass = excelColumnStyle.excelCellStyleClass();
        // 1. Case of Enum
        if (excelCellStyleClass.isEnum()) {
            String enumName = excelColumnStyle.enumName();
            return findExcelCellStyle(excelCellStyleClass, enumName);
        }

        // 2. Case of Class
        try {
            return excelCellStyleClass.newInstance();
        } catch (InstantiationException | IllegalAccessException e) {
            throw new RuntimeException();
        }
    }

    @SuppressWarnings("unchecked")
    private static ExcelCellStyle findExcelCellStyle(Class<?> excelCellStyles, String enumName) {
        try {
            return (ExcelCellStyle) Enum.valueOf((Class<Enum>) excelCellStyles, enumName);
        } catch (NullPointerException e) {
            throw new NullPointerException();
        } catch (IllegalArgumentException e) {
            throw new IllegalArgumentException();
        }
    }

}

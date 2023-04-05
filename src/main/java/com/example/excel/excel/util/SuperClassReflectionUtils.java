package com.example.excel.excel.util;

import lombok.NoArgsConstructor;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.File;
import java.io.IOException;
import java.lang.annotation.Annotation;
import java.lang.reflect.Field;
import java.util.*;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

@NoArgsConstructor
public final class SuperClassReflectionUtils {

    public static List<Field> getAllFields(Class<?> clazz) {
        List<Field> fields = new ArrayList<>();
        for (Class<?> clazzInClasses : getAllClassesIncludingSuperClasses(clazz, true)) {
            fields.addAll(Arrays.asList(clazzInClasses.getDeclaredFields()));
        }
        return fields;
    }

    public static Annotation getAnnotation(Class<?> clazz,
                                           Class<? extends Annotation> targetAnnotation) {
        for (Class<?> clazzInClasses : getAllClassesIncludingSuperClasses(clazz, false)) {
            if (clazzInClasses.isAnnotationPresent(targetAnnotation)) {
                return clazzInClasses.getAnnotation(targetAnnotation);
            }
        }
        return null;
    }

    public static Field getField(Class<?> clazz, String name) throws Exception {
        for (Class<?> clazzInClasses : getAllClassesIncludingSuperClasses(clazz, false)) {
            for (Field field : clazzInClasses.getDeclaredFields()) {
                if (field.getName().equals(name)) {
                    return clazzInClasses.getDeclaredField(name);
                }
            }
        }
        throw new NoSuchFieldException();
    }

    private static List<Class<?>> getAllClassesIncludingSuperClasses(Class<?> clazz, boolean fromSuper) {
        List<Class<?>> classes = new ArrayList<>();
        while (clazz != null) {
            classes.add(clazz);
            clazz = clazz.getSuperclass();
        }
        if (fromSuper) {
            Collections.reverse(classes);
        }
        return classes;
    }

    public static List<String> getHeader(final File file, Integer rownum) {

        // rownum 이 입력되지 않으면 default로 0 번째 라인을 header 로 판단
        if (Objects.isNull(rownum)) rownum = 0;

        final Workbook workbook;
        try {
            workbook = WorkbookFactory.create(file);
        } catch (IOException e) {
            throw new IllegalArgumentException(e);
        }
        // 시트 수 (첫번째에만 존재시 0)
        final Sheet sheet = workbook.getSheetAt(0);

        // 타이틀 가져오기
        Row title = sheet.getRow(rownum);

        try {
            workbook.close();
        } catch (IOException e) {
            throw new IllegalArgumentException(e);
        }

        return IntStream.range(0, title.getPhysicalNumberOfCells())
                .mapToObj(cellIndex -> title.getCell(cellIndex).getStringCellValue())
                .collect(Collectors.toList());
    }

}

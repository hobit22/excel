package com.example.excel.excel.util;

import com.example.excel.excel.ExcelSheet;
import com.example.excel.excel.error.ExcelReaderError;
import com.example.excel.excel.error.ExcelReaderErrorField;
import com.example.excel.excel.reader.ExcelReader;
import org.apache.commons.lang3.exception.ExceptionUtils;
import org.apache.poi.ss.formula.eval.ErrorEval;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.springframework.util.StringUtils;

import javax.validation.ConstraintViolation;
import javax.validation.Validation;
import javax.validation.Validator;
import javax.validation.ValidatorFactory;
import java.lang.reflect.Field;
import java.lang.reflect.Modifier;
import java.lang.reflect.Type;
import java.util.Objects;
import java.util.Set;

public class ExcelUtils {
    /**
     * Cell 데이터를 Type 별로 체크 하여 String 데이터로 변환함
     * String 데이터로 우선 변환해야 함
     *
     * @param cell 요청 엑셀 파일의 cell 데이터
     * @return String 형으로 변환된 cell 데이터
     */
    public static String getValue(Cell cell) {

        String value = null;
        // 셀 내용의 유형 판별
        switch (cell.getCellType()) {
            case STRING: // getRichStringCellValue() 메소드를 사용하여 컨텐츠를 읽음
                value = cell.getRichStringCellValue().getString();
                break;
            case NUMERIC: // 날짜 또는 숫자를 포함 할 수 있으며 아래와 같이 읽음
                if (DateUtil.isCellDateFormatted(cell))
                    value = cell.getLocalDateTimeCellValue().toString();
                else
                    value = String.valueOf(cell.getNumericCellValue());
                if (value.endsWith(".0"))
                    value = value.substring(0, value.length() - 2);
                break;
            case BOOLEAN:
                value = String.valueOf(cell.getBooleanCellValue());
                break;
            case FORMULA:
                value = String.valueOf(cell.getCellFormula());
                break;
            case ERROR:
                value = ErrorEval.getText(cell.getErrorCellValue());
                break;
            case BLANK:
            case _NONE:
            default:
                value = "";
        }
        //log.debug("## cellType = {}, value = {}",cell.getCellType(),value);
        return value;
    }

    /**
     * TypeParser 로 String으로 변환된 Cell 데이터를 객체 필드 타입에 맞게 변환하여 셋팅해줌
     *
     * @param object 요청 객체
     * @param <T>
     * @param row    엑셀 ROW 데이터
     * @return Cell 데이터를 맵핑한 오브젝트
     */
    public static <T> T setObjectMapping(T object, Row row) {

        int i = 0;

        if (Objects.isNull(object)) return null;

        for (Field field : SuperClassReflectionUtils.getAllFields(object.getClass())) {

            if (Modifier.isPrivate(field.getModifiers())) continue;
            ExcelSheet annotation = object.getClass().getAnnotation(ExcelSheet.class);
            String sheetName = annotation.sheetName();

            field.setAccessible(true);
            String cellValue = null;

            try {
                if (i < row.getPhysicalNumberOfCells()) { //유효한 Cell 영역 까지만
                    cellValue = ExcelUtils.getValue(row.getCell(i));
                    Object setData = null;
                    if (!StringUtils.isEmpty(cellValue)) {
                        setData = parseType(cellValue, field.getType());
                    }
                    field.set(object, setData);
                    checkValidation(object, row, i, cellValue, field.getName());
                }
            } catch (NumberFormatException e) {
                ExcelReaderError error = ExcelReaderError.TYPE;
                ExcelReader.errorFieldList.add(ExcelReaderErrorField.builder()
                        .sheetName(sheetName)
                        .type(error.name())
                        .row(row.getRowNum() + 1)
                        .field(field.getName())
                        .fieldHeader(ExcelReader.headerList.get(i))
                        .inputData(cellValue)
                        .message(error.getMessage() +
                                "데이터 필드타입 - " + field.getType().getSimpleName() +
                                ", 입력값 필드타입 - " + cellValue.getClass().getSimpleName())
                        .exceptionMessage(ExceptionUtils.getRootCauseMessage(e))
                        .build());
            } catch (Exception e) {
                ExcelReaderError error = ExcelReaderError.UNKNOWN;
                ExcelReader.errorFieldList.add(ExcelReaderErrorField.builder()
                        .sheetName(sheetName)
                        .type(error.name())
                        .row(row.getRowNum() + 1)
                        .field(field.getName())
                        .fieldHeader(ExcelReader.headerList.get(i))
                        .inputData(cellValue)
                        .message(error.getMessage())
                        .exceptionMessage(ExceptionUtils.getRootCauseMessage(e))
                        .build());
            }
            i++;

        }

        return object;
    }

    private static Object parseType(String input, Type targetType) {
        //ToDo 필요하다면 추후 다른 type 추가
        if (targetType.getTypeName().endsWith("String")) {
            return input;
        } else {
            return Integer.parseInt(input);
        }

    }

    /**
     * 객체에 대한 Validation 을 검증해서 검증을 통과 하지 못한 내역이 있을 경우 에러 리스트에 담는다
     *
     * @param object
     * @param row
     * @param i
     * @param <T>
     */
    private static <T> void checkValidation(T object, Row row, int i, String cellValue, String fieldName) {
        ValidatorFactory factory = Validation.buildDefaultValidatorFactory();
        Validator validator = factory.getValidator();
        Set<ConstraintViolation<T>> constraintValidations = validator.validate(object);
        ConstraintViolation<T> validData = constraintValidations.stream()
                .filter(data -> data.getPropertyPath().toString().equals(fieldName))
                .findFirst().orElse(null);

        if (Objects.isNull(validData)) return;

        ExcelSheet annotation = object.getClass().getAnnotation(ExcelSheet.class);
        String sheetName = annotation.sheetName();

        String fieldHeader = ExcelReader.headerList.get(i);
        ExcelReaderError error = ExcelReaderError.VALID;
        String exceptionMessage = validData.getMessage();

        if (validData.getMessageTemplate().contains("NotEmpty") || validData.getMessageTemplate().contains("NotNull")) {
            error = ExcelReaderError.EMPTY;
            exceptionMessage = fieldHeader + "은 필수 입력값입니다";
        }

        ExcelReader.errorFieldList.add(ExcelReaderErrorField.builder()
                .sheetName(sheetName)
                .type(error.name())
                .row(row.getRowNum() + 1)
                .field(validData.getPropertyPath().toString())
                .fieldHeader(fieldHeader)
                .inputData(cellValue)
                .message(error.getMessage())
                .exceptionMessage(exceptionMessage)
                .build());
    }
}

package com.example.excel.excel.error;

import lombok.Builder;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@Setter
@NoArgsConstructor
public class ExcelReaderErrorField {
    public String sheetName;
    public String type;
    public Integer row;
    public Integer index;
    public String field;
    public String fieldHeader;
    public String inputData;
    public String message;
    public String exceptionMessage;

    @Builder
    public ExcelReaderErrorField(String sheetName, String type, Integer row, Integer index, String field, String fieldHeader, String inputData, String message, String exceptionMessage) {
        this.sheetName = sheetName;
        this.type = type;
        this.row = row;
        this.index = index;
        this.field = field;
        this.fieldHeader = fieldHeader;
        this.inputData = inputData;
        this.message = message;
        this.exceptionMessage = exceptionMessage;
    }
}

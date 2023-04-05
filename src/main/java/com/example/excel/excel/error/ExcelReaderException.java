package com.example.excel.excel.error;

import lombok.Getter;

@Getter
public class ExcelReaderException extends RuntimeException {

    private final Object errorFieldList;

    public ExcelReaderException(Object errorFieldList) {
        this.errorFieldList = errorFieldList;
    }
}

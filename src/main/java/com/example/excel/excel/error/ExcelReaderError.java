package com.example.excel.excel.error;

import lombok.AllArgsConstructor;
import lombok.Getter;

@Getter
@AllArgsConstructor
public enum ExcelReaderError {
    TYPE("잘못된 데이터 타입: "),
    EMPTY("필수 입력값 누락"),
    VALID("유효성 검증 실패"),
    UNKNOWN("알수 없는 에러");

    private final String message;
}
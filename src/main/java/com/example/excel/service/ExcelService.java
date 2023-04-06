package com.example.excel.service;

import com.example.excel.dto.ExcelExamResultDto;
import com.example.excel.dto.ExcelUserDto;
import lombok.RequiredArgsConstructor;
import org.springframework.stereotype.Service;

import java.util.Arrays;
import java.util.List;

@Service
@RequiredArgsConstructor
public class ExcelService {

    public List<ExcelUserDto> getExcelDtos() {
        ExcelUserDto excelUserDto1 = new ExcelUserDto(1L, "홍길동", "선생님");
        ExcelUserDto excelUserDto2 = new ExcelUserDto(1L, "홍길동", "학생");

        return Arrays.asList(excelUserDto1, excelUserDto2);
    }

    public List<ExcelExamResultDto> getExcelExamResultDto() {
        ExcelExamResultDto excelUserDto1 = ExcelExamResultDto.builder()
                .lessonName("수업1")
                .ucode("23-10101")
                .username("학생A")
                .totalScore("12/16")
                .calculatedScore("73")
                .rightFlagList(Arrays.asList("맞음", "틀림", "부분점수", "미제출"))
                .build();

        return Arrays.asList(excelUserDto1);
    }
}

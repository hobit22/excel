package com.example.excel.service;

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
}

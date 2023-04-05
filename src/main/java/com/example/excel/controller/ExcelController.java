package com.example.excel.controller;


import com.example.excel.dto.ExcelExamResultDto;
import com.example.excel.dto.ExcelUserDto;
import com.example.excel.excel.file.OneSheetExcelFile;
import com.example.excel.service.ExcelService;
import lombok.RequiredArgsConstructor;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.util.List;

@RestController
@RequiredArgsConstructor
public class ExcelController {
    private final ExcelService excelService;

    @GetMapping("/excel/user")
    public void downloadExcel(HttpServletResponse response) throws IOException {
        List<ExcelUserDto> userDtoList = excelService.getExcelDtos();
        OneSheetExcelFile<ExcelUserDto> excelFile = new OneSheetExcelFile<>(userDtoList, ExcelUserDto.class);

        response.setHeader("Content-disposition", "attachment;filename=filename.xls");
        excelFile.write(response.getOutputStream());
    }

    @GetMapping("/excel/exam-result")
    public void downloadExamResult(HttpServletResponse response) throws IOException {
        List<ExcelExamResultDto> examResultDtoList = excelService.getExcelExamResultDto();
        OneSheetExcelFile<ExcelExamResultDto> excelFile = new OneSheetExcelFile<>(examResultDtoList, ExcelExamResultDto.class);

        response.setHeader("Content-disposition", "attachment;filename=filename.xls");
        excelFile.write(response.getOutputStream());
    }

//    @PostMapping("/excel/upload")
//    public ApiResponse uploadExcelTeacher(@RequestPart MultipartFile file, Principal principal) {
//        excelFacade.uploadExcelTeacher(file, Long.parseLong(principal.getName()));
//        return ApiResponse.noContent();
//    }
}
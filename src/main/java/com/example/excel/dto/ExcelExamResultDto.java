package com.example.excel.dto;


import com.example.excel.excel.style.DefaultExcelCellStyle;
import com.example.excel.excel.style.DefaultHeaderStyle;
import com.example.excel.excel.style.ExcelColumn;
import com.example.excel.excel.style.ExcelColumnStyle;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;

import java.util.List;


@DefaultHeaderStyle(style = @ExcelColumnStyle(excelCellStyleClass = DefaultExcelCellStyle.class))
@Builder
public class ExcelExamResultDto {
    @ExcelColumn(headerName = "수업명")
    public String lessonName;
    @ExcelColumn(headerName = "학번")
    public String ucode;
    @ExcelColumn(headerName = "이름")
    public String username;
    @ExcelColumn(headerName = "총점")
    public String totalScore;
    @ExcelColumn(headerName = "백점환산")
    public String calculatedScore;
    @ExcelColumn(headerName = "문항별 정오")
    public List<String> rightFlagList;
}

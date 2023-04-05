package com.example.excel.dto;


import com.example.excel.excel.style.DefaultExcelCellStyle;
import com.example.excel.excel.style.DefaultHeaderStyle;
import com.example.excel.excel.style.ExcelColumn;
import com.example.excel.excel.style.ExcelColumnStyle;
import lombok.AllArgsConstructor;

@DefaultHeaderStyle(style = @ExcelColumnStyle(excelCellStyleClass = DefaultExcelCellStyle.class))
@AllArgsConstructor
public class ExcelUserDto {
    @ExcelColumn(headerName = "id")
    public Long userId;

    @ExcelColumn(headerName = "이름")
    public String name;
    @ExcelColumn(headerName = "타입")
    public String type;
}

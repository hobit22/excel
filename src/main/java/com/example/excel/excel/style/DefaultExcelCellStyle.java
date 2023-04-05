package com.example.excel.excel.style;

import com.example.excel.excel.style.align.DefaultExcelAlign;

public class DefaultExcelCellStyle extends CustomExcelCellStyle{
    @Override
    public void configure(ExcelCellStyleConfigurer configurer) {
        configurer.foregroundColor(255, 255, 153)
                .excelAlign(DefaultExcelAlign.CENTER_CENTER);
    }
}

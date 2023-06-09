package com.example.excel.excel.style;

import org.apache.poi.ss.usermodel.CellStyle;

public abstract class CustomExcelCellStyle implements ExcelCellStyle {

    private ExcelCellStyleConfigurer configurer = new ExcelCellStyleConfigurer();

    public CustomExcelCellStyle() {
        configure(configurer);
    }

    public abstract void configure(ExcelCellStyleConfigurer configurer);

    @Override
    public void apply(CellStyle cellStyle) {
        configurer.configure(cellStyle);
    }

}
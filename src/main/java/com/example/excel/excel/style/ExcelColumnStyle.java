package com.example.excel.excel.style;


public @interface ExcelColumnStyle {

    Class<? extends ExcelCellStyle> excelCellStyleClass();

    String enumName() default "";

}

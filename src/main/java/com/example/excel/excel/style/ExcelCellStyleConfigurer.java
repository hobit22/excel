package com.example.excel.excel.style;


import com.example.excel.excel.style.align.ExcelAlign;
import com.example.excel.excel.style.align.NoExcelAlign;
import com.example.excel.excel.style.color.DefaultExcelColor;
import com.example.excel.excel.style.color.ExcelColor;
import com.example.excel.excel.style.color.NoExcelColor;
import org.apache.poi.ss.usermodel.CellStyle;

public class ExcelCellStyleConfigurer {

    private ExcelAlign excelAlign = new NoExcelAlign();
    private ExcelColor foregroundColor = new NoExcelColor();
//    private ExcelBorders excelBorders = new NoExcelBorders();


    public ExcelCellStyleConfigurer excelAlign(ExcelAlign excelAlign) {
        this.excelAlign = excelAlign;
        return this;
    }

    public ExcelCellStyleConfigurer foregroundColor(int red, int blue, int green) {
        this.foregroundColor = DefaultExcelColor.rgb(red, blue, green);
        return this;
    }

    public void configure(CellStyle cellStyle) {
        excelAlign.apply(cellStyle);
        foregroundColor.applyForeground(cellStyle);
//        excelBorders.apply(cellStyle);
    }
}

package com.example.excel.excel.file;

import com.example.excel.excel.resource.*;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.util.Collections;
import java.util.List;

import static com.example.excel.excel.util.SuperClassReflectionUtils.getField;

public abstract class SXSSFExcelFile<T> implements ExcelFile<T> {

    protected static final SpreadsheetVersion supplyExcelVersion = SpreadsheetVersion.EXCEL2007;

    protected SXSSFWorkbook wb;
    protected SXSSFSheet sheet;
    protected ExcelRenderResource resource;

    /**
     * SXSSFExcelFile
     *
     * @param type Class type to be rendered
     */
    public SXSSFExcelFile(Class<T> type) {
        this(Collections.emptyList(), type, new DefaultDataFormatDecider());
    }

    /**
     * SXSSFExcelFile
     *
     * @param data List Data to render excel file. data should have at least one @ExcelColumn on fields
     * @param type Class type to be rendered
     */
    public SXSSFExcelFile(List<T> data, Class<T> type) {
        this(data, type, new DefaultDataFormatDecider());
    }

    /**
     * SXSSFExcelFile
     *
     * @param data              List Data to render excel file. data should have at least one @ExcelColumn on fields
     * @param type              Class type to be rendered
     * @param dataFormatDecider Custom DataFormatDecider
     */
    public SXSSFExcelFile(List<T> data, Class<T> type, DataFormatDecider dataFormatDecider) {
        validateData(data);
        this.wb = new SXSSFWorkbook();
        ExcelRenderResource resource = ExcelRenderResourceFactory.prepareRenderResource(type, wb, dataFormatDecider, data);
        this.resource = resource;
        renderExcel(data);
    }

    protected void validateData(List<T> data) {
    }

    protected abstract void renderExcel(List<T> data);

    protected void renderHeadersWithNewSheet(Sheet sheet, int rowIndex, int columnStartIndex) {
        Row row = sheet.createRow(rowIndex);
        int columnIndex = columnStartIndex;
        for (String dataFieldName : resource.getDataFieldNames()) {
            if (resource.getExcelHeaderName(dataFieldName) instanceof List){
                List excelHeaderName = (List) resource.getExcelHeaderName(dataFieldName);

                for (Object o : excelHeaderName) {
                    Cell cell = row.createCell(columnIndex++);
                    cell.setCellStyle(resource.getCellStyle(dataFieldName, ExcelRenderLocation.HEADER));
                    cell.setCellValue((String) o);
                }

            } else {
                Cell cell = row.createCell(columnIndex++);
                cell.setCellStyle(resource.getCellStyle(dataFieldName, ExcelRenderLocation.HEADER));
                cell.setCellValue((String) resource.getExcelHeaderName(dataFieldName));
            }
        }
    }

    protected void renderBody(Object data, int rowIndex, int columnStartIndex) {
        Row row = sheet.createRow(rowIndex);
        int columnIndex = columnStartIndex;
        for (String dataFieldName : resource.getDataFieldNames()) {
            try {
                Field field = getField(data.getClass(), (dataFieldName));
                field.setAccessible(true);

                Object fieldData = field.get(data);
                if (field.getType().equals(List.class)) {
                    List<?> listData = (List<?>) fieldData;
                    for (Object element : listData) {
                        Cell cell = row.createCell(columnIndex++);
                        cell.setCellStyle(resource.getCellStyle(dataFieldName, ExcelRenderLocation.BODY));
                        renderCellValue(cell, element);
                    }
                } else {
                    Cell cell = row.createCell(columnIndex++);
                    cell.setCellStyle(resource.getCellStyle(dataFieldName, ExcelRenderLocation.BODY));
                    renderCellValue(cell, fieldData);
                }

            } catch (Exception e) {
//            throw new CustomException(EXCEL_RENDER_FAILED);
            }
        }

    }

    private void renderCellValue(Cell cell, Object cellValue) {
        if (cellValue instanceof Number) {
            Number numberValue = (Number) cellValue;
            cell.setCellValue(numberValue.doubleValue());
            return;
        }
        cell.setCellValue(cellValue == null ? "-" : cellValue.toString());
    }

    public void write(OutputStream stream) throws IOException {

        sheet.trackAllColumnsForAutoSizing();

        short lastCellNum = sheet.getRow(sheet.getLastRowNum()).getLastCellNum();
        for (int c = 0; c < lastCellNum; c++) {
            sheet.autoSizeColumn(c);
            sheet.setColumnWidth(c, (sheet.getColumnWidth(c)) + 1024);
        }

        wb.write(stream);
        wb.close();
        wb.dispose();
        stream.close();
    }

}

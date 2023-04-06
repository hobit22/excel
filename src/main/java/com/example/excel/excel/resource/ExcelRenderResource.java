package com.example.excel.excel.resource;

import lombok.AllArgsConstructor;
import lombok.Getter;
import org.apache.poi.ss.usermodel.CellStyle;

import java.util.List;
import java.util.Map;

@AllArgsConstructor
@Getter
public class ExcelRenderResource {

    private PreCalculatedCellStyleMap styleMap;

    private Map<String, Object> excelHeaderNames;
    private List<String> dataFieldNames;

    public CellStyle getCellStyle(String dataFieldName, ExcelRenderLocation excelRenderLocation) {
        return styleMap.get(ExcelCellKey.of(dataFieldName, excelRenderLocation));
    }

    public Object getExcelHeaderName(String dataFieldName) {
        return excelHeaderNames.get(dataFieldName);
    }

}

package com.example.excel.excel;

import lombok.Data;
import org.apache.poi.ss.usermodel.Row;

@Data
public abstract class ExcelAbstractSheet {

    public Integer index;

    public abstract <T extends ExcelAbstractSheet> T from(Row row);
}

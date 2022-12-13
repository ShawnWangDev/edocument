package pers.wangsc.edocument.excel;

import org.apache.poi.ss.usermodel.Row;

public interface ExcelReader {
    int getLastRowNumber();
    Row getRow(int rowNo);
}

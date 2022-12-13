package pers.wangsc.edocument.excel;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

public class XlsxReader implements ExcelReader {

    private XSSFWorkbook workbook;
    private XSSFSheet sheet;

    public XlsxReader(String filePath) {
        File excelFile = new File(filePath);
        if (excelFile.exists()) {
            try {
                InputStream inputStream = new FileInputStream(excelFile);
                workbook = new XSSFWorkbook(inputStream);
                sheet=workbook.getSheetAt(0);
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    public XlsxReader(String filePath,int sheetIndex){
        this(filePath);
        sheet=workbook.getSheetAt(sheetIndex);
    }

    @Override
    public int getLastRowNumber() {
        return sheet.getLastRowNum();
    }

    @Override
    public Row getRow(int rowIndex) {
        return sheet.getRow(rowIndex);
    }
}

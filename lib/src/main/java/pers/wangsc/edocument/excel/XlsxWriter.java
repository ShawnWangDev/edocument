package pers.wangsc.edocument.excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.List;

public class XlsxWriter implements ExcelWriter {
    private final File excelFile;
    private XSSFWorkbook workbook;
    private int sheetAt;
    private XSSFSheet sheet;

    public XlsxWriter(String filePath) {
        this.excelFile = new File(filePath);
        this.sheetAt = 0;
        if (excelFile.exists()) {
            try {
                InputStream inputStream = new FileInputStream(excelFile);
                this.workbook = new XSSFWorkbook(inputStream);
                inputStream.close();
            } catch (IOException e) {
                throw new RuntimeException(e);
            }
            this.sheet = workbook.getSheetAt(sheetAt);
        }
    }

    public XlsxWriter(String filePath, int sheetAt) {
        this(filePath);
        this.sheetAt = sheetAt;
    }

    public Boolean shiftRows(int rowAt, int n) {
        int lastRowNum = sheet.getLastRowNum();
        if (lastRowNum > -1 && lastRowNum >= rowAt) {
            sheet.shiftRows(rowAt, lastRowNum, n);
            return true;
        }
        return false;
    }

    @Override
    public void writeGrid(Grid grid) {
        XSSFSheet sheet = this.workbook.getSheetAt(this.sheetAt);
        Row row = sheet.createRow(grid.getRowAt());
        Cell cell = row.createCell(grid.getColAt());
        cell.setCellValue(String.valueOf(grid.getData()));
    }

    @Override
    public void writeLine(Line line) {
        shiftRows(line.getRowAt(), 1);
        Row row = sheet.createRow(line.getRowAt());
        for (var entry : line.getData().entrySet()) {
            Cell cell = row.createCell(entry.getKey());
            cell.setCellValue(String.valueOf(entry.getValue()));
        }

    }

    @Override
    public void writeLines(int rowAt, List<Line> linesList) {
        int linesSize = linesList.size();
        shiftRows(rowAt, linesSize);
        for (int i = 0; i < linesSize; i++) {
            Row row = sheet.createRow(rowAt + i);
            for (var entry : linesList.get(i).getData().entrySet()) {
                row.createCell(entry.getKey()).setCellValue(String.valueOf(entry.getValue()));
            }
        }
    }

    @Override
    public void save() {
        try {
            FileOutputStream out = new FileOutputStream(excelFile);
            workbook.write(out);
            workbook.close();
            out.close();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }
}

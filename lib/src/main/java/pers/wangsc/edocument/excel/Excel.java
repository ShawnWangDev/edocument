package pers.wangsc.edocument.excel;

import org.apache.poi.ss.usermodel.Row;
import pers.wangsc.edocument.util.FileType;

import java.util.List;

public class Excel implements ExcelWriter, ExcelReader {
    private ExcelWriter excelWriter;
    private ExcelReader excelReader;
    private Boolean valid;
    public Excel(String filePath) {
        String fileType = FileType.judge(filePath);
        assert fileType != null;
        if (fileType.equals("Office 2007")) {
            excelWriter = new XlsxWriter(filePath);
            excelReader = new XlsxReader(filePath);
        }
        valid = excelReader != null;
    }

    public Boolean getValid() {
        return valid;
    }

    @Override
    public Boolean shiftRows(int rowAt, int n) {
        return excelWriter.shiftRows(rowAt, n);
    }

    @Override
    public void writeGrid(Grid grid) {
        excelWriter.writeGrid(grid);
    }

    @Override
    public void writeLine(Line line) {
        excelWriter.writeLine(line);
    }

    @Override
    public void writeLines(int rowAt, List<Line> linesList) {
        excelWriter.writeLines(rowAt, linesList);
    }

    @Override
    public void save() {
        excelWriter.save();
    }

    @Override
    public int getLastRowNumber() {
        return excelReader.getLastRowNumber();
    }

    @Override
    public Row getRow(int rowNo) {
        return excelReader.getRow(rowNo);
    }
}

package pers.wangsc.edocument.excel;

import java.util.List;

public interface ExcelWriter {
    Boolean shiftRows(int rowAt, int n);
    void writeGrid(Grid grid);
    void writeLine(Line line);
    void writeLines(int rowAt, List<Line> linesList);
    void save();
}

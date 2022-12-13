/*
 * This Java source file was generated by the Gradle 'init' task.
 */
package pers.wangsc.edocument;

import org.junit.jupiter.api.Test;
import pers.wangsc.edocument.excel.Excel;
import pers.wangsc.edocument.excel.Line;

import java.util.Arrays;

class ExcelTest {

    String excelPath="D:\\_test\\edocument\\test.xlsx";

    @Test void writeLine(){
        Excel excel=new Excel(excelPath);
        excel.writeLine(new Line(Arrays.stream(new String[]{"a","b","gamma"}).toList()));
        excel.save();
    }

    @Test void writeLineWithRowNo(){
        Excel excel=new Excel(excelPath);
        excel.writeLine(new Line(5,Arrays.stream(new String[]{"1","2","3"}).toList()));
        excel.save();
    }

    @Test void getLastRowNo(){
        Excel excel=new Excel(excelPath);
        System.out.println("\n-----\nLast Row Number:\n"+excel.getLastRowNumber());
    }
}
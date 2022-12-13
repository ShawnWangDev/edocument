package pers.wangsc.edocument.word;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.hwpf.usermodel.Table;
import org.apache.poi.hwpf.usermodel.TableIterator;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

public class DocReader implements WordReader {
    private HWPFDocument hwpfDocument;

    public DocReader(String filePath) {
        File wordFile = new File(filePath);
        if (wordFile.exists()) {
            try {
                InputStream inputStream = new FileInputStream(wordFile);
                hwpfDocument = new HWPFDocument(inputStream);
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    private List<Table> getTableList() {
        List<Table> tableList = new ArrayList<>();
        Range range = hwpfDocument.getRange();
        TableIterator tableIterator = new TableIterator(range);
        while (tableIterator.hasNext()) {
            Table t = tableIterator.next();
            tableList.add(t);
        }
        return tableList;
    }

    @Override
    public List<String> getTableContent(int tableAt, int[] locations) {
        Table table = getTableList().get(tableAt);
        List<String> data = new ArrayList<>();
        for (int i = 0; i < locations.length - 1; i += 2) {
            data.add(table.getRow(locations[i]).getCell(locations[i + 1]).text().trim());
        }
        return data;
    }
}

package pers.wangsc.edocument.word;

import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;

import java.io.*;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Objects;

public class DocxReader implements WordReader {
    private XWPFDocument xwpfDocument;

    public DocxReader(String filePath) {
        File wordFile = new File(filePath);
        if (wordFile.exists()) {
            try {
                InputStream inputStream = new FileInputStream(wordFile);
                xwpfDocument = new XWPFDocument(inputStream);
            } catch (IOException e) {
                throw new RuntimeException(e);
            }
        }
    }

    private List<XWPFTable> getTableList() {
        Iterator<IBodyElement> docElementsIterator = this.xwpfDocument.getBodyElementsIterator();
        while (docElementsIterator.hasNext()) {
            IBodyElement docElement = docElementsIterator.next();
            if ("TABLE".equalsIgnoreCase(docElement.getElementType().name())) {
                return docElement.getBody().getTables();
            }
        }
        return null;
    }

    @Override
    public List<String> getTableContent(int tableAt, int[] locations) {
        XWPFTable table = Objects.requireNonNull(getTableList()).get(tableAt);
        List<String> data = new ArrayList<>();
        for (int i = 0; i < locations.length - 1; i += 2) {
            data.add(table.getRow(locations[i]).getCell(locations[i + 1]).getText().trim());
        }
        return data;
    }
}

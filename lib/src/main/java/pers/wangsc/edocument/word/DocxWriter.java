package pers.wangsc.edocument.word;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.io.*;
import java.util.Map;

public class DocxWriter implements WordWriter {
    private File wordFile;
    private XWPFDocument xwpfDocument;

    public DocxWriter(String filePath) {
        wordFile = new File(filePath);
        if (wordFile.exists()) {
            try {
                InputStream inputStream = new FileInputStream(wordFile);
                xwpfDocument = new XWPFDocument(inputStream);
            } catch (IOException e) {
                throw new RuntimeException(e);
            }
        }
    }

    @Override
    public void writeParagraph(String text, String fontFamily, Double fontSize) {
        XWPFParagraph paragraph = this.xwpfDocument.createParagraph();
        XWPFRun run = paragraph.createRun();
        run.setText(text);
        run.setFontFamily(fontFamily);
        run.setFontSize(fontSize);
    }

    @Override
    public void replaceParagraphText(Map<String, String> map) {
        for (var paragraph : this.xwpfDocument.getParagraphs()) {
            for (var run : paragraph.getRuns()) {
                String text = run.getText(run.getTextPosition());
                for (var entry : map.entrySet()) {
                    text = text.replace(entry.getKey(), entry.getValue());
                }
                run.setText(text, 0);
            }
        }
    }

    @Override
    public void save() {
        try {
            if (!wordFile.exists()) {
                wordFile.createNewFile();
            }
            FileOutputStream out = new FileOutputStream(wordFile);
            xwpfDocument.write(out);
            xwpfDocument.close();
            out.close();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }
}
